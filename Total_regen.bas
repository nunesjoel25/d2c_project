Attribute VB_Name = "Total_regen"
Sub copyRange_total()
    Dim total_regen As New frontendRegeneration
    Dim copyRange_total As Range
    Dim pasteRange_total As Range
    Dim combinations_total As Long
    Set total_regen.copyRange_total = Sheets("TotalTemplateView").Range("D6:AC41")
    Set total_regen.pasteRange_total = Sheets("Total").Range("D6")
     total_regen.combinations_total = Sheets("ph3_combinations").Range("B2")
    
        total_regen.CopyPaste total_regen.copyRange_total, total_regen.pasteRange_total, total_regen.combinations_total

End Sub
Sub MaterialListcopy_total()

    
    Set sourceSheet = Worksheets("total_list")
    Set targetSheet = Worksheets("Total")
    Set sourceTable = sourceSheet.ListObjects("total_list")
    Set targetCell = targetSheet.Range("B6")
    'Copy the table to the target sheet
     sourceTable.DataBodyRange.Copy Destination:=targetCell
    
End Sub

Sub clear_total()
Call clear_outline
Sheets("Total").Unprotect Password:="clausus"
Sheets("Total").Range("b6:AC10000").clear
Sheets("Total").Range("B6:C10000").clear

End Sub
Sub clear_outline()
Worksheets("Total").Cells.ClearOutline
End Sub
Sub HighlightTextWithAsterisks_total()
  Worksheets("Total").Activate
    Dim lastRow As Long
    Dim columnToSearch As Long
    Dim cell As Range
    Dim targetSheet As Worksheet
    
    
    columnToSearch = 7
    
   
    Set targetSheet = ThisWorkbook.Sheets("Total")
    
    lastRow = targetSheet.Cells(targetSheet.Rows.count, columnToSearch).End(xlUp).Row
    
 
    For Each cell In targetSheet.Range("G6:G" & lastRow)
        
        
        If InStr(cell.value, "^") > 0 Then
            
            targetSheet.Range("B" & cell.Row & ":C" & cell.Row).Font.Color = RGB(174, 170, 170)
            targetSheet.Range("B" & cell.Row & ":AC" & cell.Row).Interior.Color = RGB(174, 170, 170)
            
        End If
        
    Next cell
    
End Sub

Sub Hide_Columns_Containing_Value_total()
'
    Rows("1:3").Hidden = True
 Worksheets("Total").Activate
Dim c As Range

    For Each c In Range("A2:G2").Cells
        If c.value = "hide" Then
            c.EntireColumn.Hidden = True
        End If
    Next c

End Sub
Sub refresh_Pivot()
If Sheets("User Selections").Range("G7").value = "Key" Then
        Sheets("Total_Pivot").PivotTables("PivotTable1").RefreshTable
    Else
        'Sheets("Non_key_Pivot").PivotTables("PivotTable2").RefreshTable
    End If

End Sub
Sub Unhide_Columns_Containing_Value_total()
'
 Worksheets("Total").Activate
 Rows("1:3").Hidden = False
Dim c As Range

    For Each c In Range("D2:H2").Cells
        If c.value = "hide" Then
            c.EntireColumn.Hidden = False
        End If
    Next c

End Sub
Sub GroupRowsWithAsterisk_total()
   Worksheets("Total").Activate
    Dim lastRow As Long
    Dim cell As Range
    Dim rng As Range
    
    lastRow = Cells(Rows.count, "I").End(xlUp).Row
    
    For Each cell In Range("G1:G" & lastRow)
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
        Range(rng, Cells(lastRow, "G")).Rows.Group
    End If
    
      ActiveSheet.Outline.ShowLevels RowLevels:=1, ColumnLevels:=1

End Sub
Sub calculate_pivot_sheet()
Set ws1 = ThisWorkbook.Worksheets("Non_key_Pivot")
Set ws2 = ThisWorkbook.Worksheets("Total_Pivot")
If Sheets("User Selections").Range("G7").value = "Key" Then
    ws2.Calculate
    Else
    ws1.Calculate
    
End If
End Sub
