VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "frontendRegeneration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
    Public copyRange_input As Range
    Public pasteRange_input As Range
    Public copyRange_non_key As Range
    Public pasteRange_non_key As Range
    Public pasteRange_koro As Range
    Public combinations_input As Long
    Public combinations_koro As Long
    Private m_UserSelectionLevel As String
    Private m_UserSelection_indicator As String
    Public copyRange_koro As Range
    Public copyRange_total As Range
    Public pasteRange_total As Range
    Public combinations_total As Long
    Public sourceSheet_traffic As Worksheet
    Public sourceSheet_orders As Worksheet
    Public targetSheet_traffic As Worksheet
    Public targetSheet_orders As Worksheet
    Public sourceTable_traffic As ListObject
    Public sourceTable_quantity As ListObject
    Public targetCell_traffic As Range
    Public targetCell_orders As Range
 
    
Property Get NonKeyClearRange() As Range
NonKeyClearRange = m_NonKeyClearRange
End Property
Property Let NonKeyClearRange(clear_range As Range)
m_NonKeyClearRange = clear_range
End Property
Property Get UserSelectionLevel() As String
UserSelectionLevel = m_UserSelectionLevel
End Property
Property Let UserSelectionLevel(value As String)
m_UserSelectionLevel = value
End Property
Property Get UserSelectionIndicator() As String
UserSelectionIndicator = m_UserSelection_indicator
End Property

Public Sub Class_Initialize()
m_UserSelectionLevel = UserSelections.Range("C6").value
m_UserSelection_indicator = UserSelections.Range("G7").value
End Sub
Public Sub CopyPaste(copyRange As Range, pasteRange As Range, count As Long)
    Dim i As Long
    
    For i = 1 To count
        copyRange.Copy Destination:=pasteRange
        Set pasteRange = pasteRange.offSet(copyRange.Rows.count, 0)
    Next i
End Sub

Public Sub MaterialListcopy(source_list As ListObject, target_list As ListObject)
    
    Dim sourceSheet As Worksheet
    Dim targetSheet As Worksheet
    Dim sourceTable As ListObject

    sourceTable.DataBodyRange.Copy Destination:=targetCell
  
End Sub
Public Sub groupRows(rng_grp As Range, Optional ws As Worksheet)
    Dim lastRow As Long
    Dim cell As Range
    Dim rng As Range
    
    lastRow = ws.Cells(Rows.count, "C").End(xlUp).Row

    For Each cell In rng_grp
        If InStr(1, cell.value, "*", vbTextCompare) > 0 Then
            If rng Is Nothing Then
                Set rng = cell
            Else
                Set rng = Union(rng, cell)
            End If
        Else
            If Not rng Is Nothing Then
                Range(rng, cell.offSet(-1, 0)).Rows.Group
                Set rng = Nothing
            End If
        End If
    Next cell
    
    If Not rng Is Nothing Then
        Range(rng, Cells(lastRow, "F")).Rows.Group
    End If
    
    ActiveSheet.Outline.ShowLevels RowLevels:=1, ColumnLevels:=1
End Sub

Public Sub groupColumns()
    Dim ws As Worksheet
    Dim rangeToCheck As Range
    Dim cell As Range
    Dim startColumn As Long
    Dim endColumn As Long
    Dim groupingInProgress As Boolean
    
  
    Set ws = ThisWorkbook.Sheets("Input Sheet")
    
   
    Set rangeToCheck = ws.Range("C4:AW4")
    
    For Each cell In rangeToCheck
        If cell.value = "colapse" Then
            If Not groupingInProgress Then
                startColumn = cell.Column
                groupingInProgress = True
            End If
        Else
            If groupingInProgress And endColumn = 0 Then
                endColumn = cell.Column - 1
                
                ws.Range(ws.Cells(1, startColumn), ws.Cells(1, endColumn)).EntireColumn.Group
                groupingInProgress = False
                startColumn = 0
                endColumn = 0
            End If
        End If
    Next cell
    
  
    If groupingInProgress And endColumn = 0 Then
        endColumn = rangeToCheck.Columns.combinations_input
        
        ws.Range(ws.Cells(1, startColumn), ws.Cells(1, endColumn)).EntireColumn.Group
    End If
End Sub

Public Sub u_hide_columns_rows(ws As Worksheet, bool As Boolean)
    Dim i As Range
    Dim lastRow As Long

    Dim c As Range
    
    lastRow = ws.Cells(ws.Rows.count, "F").End(xlUp).Row
    For Each i In ws.Range("B1:B8").Cells
        If i.value = "hide" Then
            i.EntireRow.Hidden = bool
        End If
    Next i
    
    For Each c In ws.Range("A5:I5").Cells
        If c.value = "hide" Then
            c.EntireColumn.Hidden = bool
        End If
    Next c
End Sub

Public Sub formatrows(ws As Worksheet, column_number As Long)
    ws.Activate
    
    Dim lastRow As Long
    Dim columnToSearch As Long
    Dim cell As Range
    Dim targetSheet_koro As Worksheet
    
    columnToSearch = column_number
    
    Set targetSheet_koro = ws
    lastRow = targetSheet_koro.Cells(targetSheet_koro.Rows.count, columnToSearch).End(xlUp).Row
    
    For Each cell In targetSheet_koro.Range("F1:I" & lastRow)
        If InStr(cell.value, "^") > 0 Then
            targetSheet_koro.Range("C" & cell.Row & ":H" & cell.Row).Font.Color = RGB(174, 170, 170)
            targetSheet_koro.Range("C" & cell.Row & ":AD" & cell.Row).Interior.Color = RGB(174, 170, 170)
        End If
    Next cell
End Sub

Public Sub clearAll(namedRange As Range, ws As Worksheet, offset_range As Long)
        namedRange.offSet(offset_range, 0).clear
    With ws
       .Cells.ClearOutline
        ActiveWindow.FreezePanes = False
    End With
    
End Sub

Public Sub RefreshQueries(ws As Worksheet, object_name As ListObject)
 ActiveWorkbook.Worksheets(ws).ListObjects(object_name) _
    .QueryTable.Refresh BackgroundQuery:=False
End Sub

Public Sub FreezePanes(rng As Range)

    With ActiveWindow
        .FreezePanes = False
        .ScrollRow = 1
        .ScrollColumn = 1
         rng.Select
        .FreezePanes = True
    End With
    rng.Activate
End Sub
Public Function ToggleSheetVisibility(wWs_arr() As Variant) As String
    Dim ws As Worksheet
    Dim foundSheet As Boolean
    Dim w As Variant
    For Each ws In Application.ActiveWorkbook.Worksheets
        foundSheet = False
        For Each w In wWs_arr
            If w Is ws Then
                foundSheet = True
                Exit For
            End If
        Next w
        If Not foundSheet Then
            ws.Visible = xlSheetVeryHidden
        End If
    Next ws

    ToggleSheetVisibility = "Sheets hidden successfully."
End Function

Public Sub turn_very_hidden(ws1 As Worksheet, ws2 As Worksheet, Optional ws3 As Worksheet)
    Dim SheetsArray() As Variant
    
    SheetsArray = Array(ws1, ws2, ws3)
 
    ToggleSheetVisibility SheetsArray
End Sub
Public Sub QueryRefresh(columnNumber As Double)
    Dim sheetId As New collection
    Dim queryId As New collection
    Dim lastRow As Long
    Dim ws As Worksheet
    
    Set ws = Worksheets("Settings")
       
    lastRow = ws.Cells(Rows.count, columnNumber).End(xlUp).Row
    Dim i As Integer
    
    For i = 28 To lastRow
        sheetId.Add Settings.Cells(i, columnNumber).value
        queryId.Add Settings.Cells(i, columnNumber + 1).value
    Next i
    
    Dim n As Integer
    For n = 1 To sheetId.count
        ActiveWorkbook.Worksheets(sheetId(n)).ListObjects(queryId(n)).QueryTable.Refresh BackgroundQuery:=False
    Next n
End Sub

Public Sub protect_sheet_key(sheetName As String)

'Allowing certain actions even if the worksheet is protected
Sheets(sheetName).Protect Password:="clausus", UserInterfaceOnly:=True, _
    DrawingObjects:=False, _
    Contents:=True, _
    Scenarios:=False, _
    AllowFormattingCells:=True, _
    AllowFormattingColumns:=True, _
    AllowFormattingRows:=True, _
    AllowInsertingColumns:=False, _
    AllowInsertingRows:=False, _
    AllowInsertingHyperlinks:=False, _
    AllowDeletingColumns:=False, _
    AllowDeletingRows:=False, _
    AllowSorting:=True, _
    AllowFiltering:=True, _
    AllowUsingPivotTables:=False
    Sheets(sheetName).EnableOutlining = True

End Sub
 

Public Sub unprotect_key_sheet(sheetName As String)

Sheets(sheetName).Unprotect Password:="clausus"

End Sub
 
 



