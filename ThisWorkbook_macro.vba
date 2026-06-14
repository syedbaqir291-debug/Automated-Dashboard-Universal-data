' =====================================================================
' OMAC Dashboard - Auto Pivot/Slicer Builder
' =====================================================================
' ONE-TIME SETUP (do this once to create the reusable template):
'
'   1. Open a new BLANK workbook in Excel.
'   2. Press Alt+F11 to open the VBA editor.
'   3. In the Project Explorer (left side), double-click "ThisWorkbook"
'      under "Microsoft Excel Objects" for that workbook.
'   4. Paste THE ENTIRE CONTENTS of this file into that code window.
'   5. Press Ctrl+S, choose "Excel Macro-Enabled Workbook (*.xlsm)",
'      and save it as:  dashboard_macro_template.xlsm
'   6. Place that file in the same folder as dashboard_builder.py
'      (or update MACRO_TEMPLATE_PATH in dashboard_builder.py to point
'      to it).
'
' From then on, every dashboard dashboard_builder.py generates will be
' an .xlsm built on top of this template. When the user opens it (and
' clicks "Enable Content" for macros, which Excel always asks for),
' this code runs automatically and builds real PivotTables, PivotCharts
' and cross-connected Slicers from the "Data" sheet, using the settings
' written to the hidden "Config" sheet.
'
' This macro is IDEMPOTENT - it removes anything it built last time
' before rebuilding, so re-opening the file (or opening it again after
' Baqir regenerates it) won't create duplicates.
' =====================================================================

Option Explicit

Private Const RECORD_COL As String = "__Records__"

Private Sub Workbook_Open()
    On Error GoTo ErrHandler
    BuildPivotDashboard
    Exit Sub
ErrHandler:
    MsgBox "Dashboard pivot/slicer setup ran into an issue:" & vbCrLf & vbCrLf & _
           Err.Description & vbCrLf & vbCrLf & _
           "The static charts (if any) and data are still intact.", vbExclamation, "Dashboard Setup"
End Sub

Sub BuildPivotDashboard()
    Dim wsData As Worksheet, wsCfg As Worksheet, wsDash As Worksheet, wsPiv As Worksheet
    Dim pc As PivotCache
    Dim pt As PivotTable
    Dim n As Long, i As Long, j As Long
    Dim dimNames() As String
    Dim measureMode As String, measureCol As String, measureLabel As String
    Dim srcRange As String
    Dim lastRow As Long, lastCol As Long

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Set wsData = ThisWorkbook.Sheets("Data")
    Set wsCfg = ThisWorkbook.Sheets("Config")
    Set wsDash = ThisWorkbook.Sheets("Dashboard")

    ' ---------------- read config written by dashboard_builder.py ----------------
    n = wsCfg.Cells(1, 1).Value
    If n < 1 Then GoTo Cleanup

    ReDim dimNames(1 To n)
    For i = 1 To n
        dimNames(i) = wsCfg.Cells(i + 1, 1).Value
    Next i
    measureMode = LCase(wsCfg.Cells(1, 2).Value)
    measureCol = wsCfg.Cells(2, 2).Value
    measureLabel = wsCfg.Cells(3, 2).Value

    ' ---------------- remove anything built on a previous open ----------------
    For i = wsDash.ChartObjects.Count To 1 Step -1
        wsDash.ChartObjects(i).Delete
    Next i

    Do While ThisWorkbook.SlicerCaches.Count > 0
        ThisWorkbook.SlicerCaches(1).Delete
    Loop

    If SheetExists("Live Pivots") Then
        ThisWorkbook.Sheets("Live Pivots").Delete
    End If

    ' ---------------- build one shared PivotCache from the Data table ----------------
    lastRow = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row
    lastCol = wsData.Cells(1, wsData.Columns.Count).End(xlToLeft).Column
    srcRange = "Data!" & wsData.Range(wsData.Cells(1, 1), wsData.Cells(lastRow, lastCol)).Address

    Set pc = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=srcRange)

    Set wsPiv = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsPiv.Name = "Live Pivots"

    ' ---------------- one PivotTable per dimension ----------------
    Dim ptArr(1 To 30) As PivotTable
    Dim destRow As Long
    destRow = 1

    For i = 1 To n
        Set pt = pc.CreatePivotTable(TableDestination:=wsPiv.Cells(destRow, 1), TableName:="PT_" & i)

        With pt.PivotFields(dimNames(i))
            .Orientation = xlRowField
            .Position = 1
        End With

        Select Case measureMode
            Case "count"
                pt.AddDataField pt.PivotFields(RECORD_COL), measureLabel, xlSum
            Case "sum"
                pt.AddDataField pt.PivotFields(measureCol), measureLabel, xlSum
            Case Else ' "average"
                pt.AddDataField pt.PivotFields(measureCol), measureLabel, xlAverage
        End Select

        pt.PivotFields(dimNames(i)).AutoSort xlDescending, measureLabel
        pt.RowGrand = False
        pt.ColumnGrand = False
        pt.RowAxisLayout xlTabularRow
        On Error Resume Next
        pt.TableStyle2 = "PivotStyleMedium9"
        On Error GoTo 0

        Set ptArr(i) = pt
        destRow = destRow + pt.TableRange2.Rows.Count + 3
    Next i

    ' ---------------- one PivotChart per dimension, on the Dashboard ----------------
    Dim leftPos As Double, topPos As Double
    Dim chartW As Double, chartH As Double
    chartW = 380
    chartH = 260
    topPos = 130

    For i = 1 To n
        Dim cObj As ChartObject
        If (i Mod 2) = 1 Then
            leftPos = 10
            If i > 1 Then topPos = topPos + chartH + 20
        Else
            leftPos = chartW + 30
        End If

        Set cObj = wsDash.ChartObjects.Add(leftPos, topPos, chartW, chartH)
        cObj.Chart.SetSourceData ptArr(i).TableRange2

        If ptArr(i).PivotFields(dimNames(i)).PivotItems.Count <= 5 Then
            cObj.Chart.ChartType = xlDoughnut
        Else
            cObj.Chart.ChartType = xlBarClustered
        End If

        cObj.Chart.HasTitle = True
        cObj.Chart.ChartTitle.Text = dimNames(i)
        cObj.Chart.ApplyDataLabels
    Next i

    topPos = topPos + chartH + 20

    ' ---------------- slicers, cross-connected to every PivotTable ----------------
    Dim sCache As SlicerCache
    Dim sl As Slicer
    Dim slicerLeft As Double
    slicerLeft = 10

    For i = 1 To n
        Set sCache = ThisWorkbook.SlicerCaches.Add2(ptArr(i), dimNames(i))
        For j = 1 To n
            If j <> i Then
                On Error Resume Next
                sCache.PivotTables.AddPivotTable ptArr(j)
                On Error GoTo 0
            End If
        Next j

        Set sl = sCache.Slicers.Add( _
            SlicerDestination:=wsDash, _
            Caption:=dimNames(i), _
            Name:=dimNames(i) & "_Slicer", _
            Top:=topPos, _
            Left:=slicerLeft, _
            Width:=140, _
            Height:=180)

        slicerLeft = slicerLeft + 150
    Next i

    wsDash.Activate

Cleanup:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Sub

Private Function SheetExists(sName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sName)
    On Error GoTo 0
    SheetExists = Not ws Is Nothing
End Function
