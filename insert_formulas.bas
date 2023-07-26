Attribute VB_Name = "insert_formulas"

Sub InsertFormula()
    Dim lastCol As Long, i As Long
    Dim value  As Integer
     Dim ws As Worksheet
    Dim lastRow As Long
    Dim rng As Range
    Dim cell As Range
    
    ' Set the target worksheet
    Set ws = ThisWorkbook.Sheets("Koro")
    lastRow = ws.Cells(ws.Rows.count, "J").End(xlUp).Row
    Sheets("Koro").Activate
    For i = 25 To 11 Step -1
        If Range(Cells(3, i), Cells(3, i)).value = "" Then
            lastCol = i
            Exit For
        End If
    Next i
    
    value = Sheets("Koro").Range("J3") - 1
    
Set rng = ws.Range("J2:J" & lastRow)
    
    ' Loop through each cell in the range
For Each cell In rng
    If lastCol > 0 Then 'Last column was found
        If cell.value = "Paid Search % (Input)" Or cell.value = "Email % (Input)" Or cell.value = "Social % (Input)" Then
             Range(Cells(cell.Row, lastCol - value), Cells(cell.Row, lastCol)).Formula = "=K" & cell.Row - 3
         End If
    End If
Next cell

End Sub

Sub insert_formulas_retrival()
    Dim lastCol As Long, i As Long
    Dim value  As Integer
     Dim ws As Worksheet
    Dim lastRow As Long
    Dim rng As Range
    Dim cell As Range
    
    ' Set the target worksheet
    Set ws = ThisWorkbook.Sheets("Koro") ' Update "Sheet1" to the name of your worksheet
    
    ' Find the last row in the worksheet
    lastRow = ws.Cells(ws.Rows.count, "J").End(xlUp).Row
Sheets("Koro").Activate
    'Find last column where value is "ACT" in K2:Y2
    For i = 11 To 25 'Loop backwards from column Y to column K
        If Range(Cells(3, i), Cells(3, i)).value = "*" Then
            lastCol = i 'Set last column to the column with "ACT" value
            Exit For 'Exit loop once last column is found
        End If
    Next i
    
    value = 15 - Sheets("Koro").Range("J3")
    'Insert formula in K46:Y48 range using last column
Set rng = ws.Range("J2:J" & lastRow) ' Assuming the specific values are in column A, starting from row 2
    
    ' Loop through each cell in the range
For Each cell In rng
        If cell.value = "Uplift" Or cell.value = "Paid Search % (Input)" Or cell.value = "D2C Conversion (Override)" Or cell.value = "Email % (Input)" Or cell.value = "Social % (Input)" Then
             Range(Cells(cell.Row, 11), Cells(cell.Row, 25)).Formula = "=IFNA(INDEX(Koro_live[Value],MATCH('Koro'!$I" & cell.Row & "&'Koro'!$J" & cell.Row & "&'Koro'!K$6,Koro_live[key],0)),0)"
         End If
Next cell
For Each cell In rng
        If cell.value = "Sales Quantity Override" Then
             Range(Cells(cell.Row, 11), Cells(cell.Row, 25)).Formula = "=IFNA(INDEX(ip_live[0G_QABSMG],MATCH('Koro'!$H" & cell.Row & "&'Koro'!$I" & cell.Row & "&'Koro'!K$6,ip_live[index],0)),0)"
         End If

Next cell

For i = 25 To 11 Step -1 'Loop backwards from column Y to column K
        If Range(Cells(3, i), Cells(3, i)).value = "" Then
            lastCol1 = i 'Set last column to the column with "ACT" value
            Exit For 'Exit loop once last column is found
        End If
    Next i
    
    Value1 = Sheets("Koro").Range("J3") - 1
    'Insert formula in K46:Y48 range using last column
Set rng = ws.Range("J2:J" & lastRow) ' Assuming the specific values are in column A, starting from row 2



For Each cell In rng
    If lastCol1 > 0 Then 'Last column was found
        If cell.value = "Uplift" Or cell.value = "D2C Conversion (Override)" Or cell.value = "Sales Quantity Override" Then
             Range(Cells(cell.Row, lastCol1 - Value1), Cells(cell.Row, lastCol1)).clearContents
         End If
    End If
Next cell



End Sub


Sub clearContents()
    Dim ws As Worksheet
    Dim i As Long
    Sheets("Koro").Activate
    Set ws = ThisWorkbook.Worksheets("Koro") ' Replace "Sheet1" with the name of your worksheet
    
    For i = 7 To 80 ' Loop through
        If ws.Range("J" & i).value = "Paid Search % (Input)" Or ws.Range("J" & i).value = "Uplift" Or ws.Range("J" & i).value = "D2C Conversion (Override)" Or ws.Range("J" & i).value = "Email % (Input)" Or ws.Range("J" & i).value = "Social % (Input)" Or ws.Range("J" & i).value = "Sales Quantity Override" Then
            ws.Range("K" & i & ":Y" & i).clearContents
        End If
    Next i
End Sub

Sub insert_formulas_retrival_non_key()

    Dim lastCol As Long, i As Long
    Dim value  As Integer
     Dim ws As Worksheet
    Dim lastRow As Long
    Dim rng As Range
    Dim cell As Range
    
    ' Set the target worksheet
    Set ws = ThisWorkbook.Sheets("Non-key") ' Update "Sheet1" to the name of your worksheet
    Sheets("Non-Key").Activate
    ' Find the last row in the worksheet
    lastRow = ws.Cells(ws.Rows.count, "J").End(xlUp).Row
    
    'Find last column where value is "ACT" in K2:Y2
    For i = 11 To 25 'Loop backwards from column Y to column K
        If Range(Cells(3, i), Cells(3, i)).value = "*" Then
            lastCol = i 'Set last column to the column with "ACT" value
            Exit For 'Exit loop once last column is found
        End If
    Next i
    
    'Insert formula in K46:Y48 range using last column
Set rng = ws.Range("J2:J" & lastRow) ' Assuming the specific values are in column A, starting from row 2
    
    ' Loop through each cell in the range
For Each cell In rng
        If cell.value = "D2C Conversion (Override)" Then
             Range(Cells(cell.Row, 11), Cells(cell.Row, 25)).Formula = "=IFNA(INDEX(Koro_live[Value],MATCH('Non-Key'!$I" & cell.Row & "&'Non-Key'!$J" & cell.Row & "&'Non-Key'!K$6,Koro_live[key],0)),0)"
        Else

        If cell.value = "Uplift" Then
             Range(Cells(cell.Row, 11), Cells(cell.Row, 25)).Formula = "=IFNA(INDEX(extract_ret_2[Uplift],MATCH('Non-Key'!$I" & cell.Row & "&'Non-Key'!$J" & cell.Row & "&'Non-Key'!K$6,extract_ret_2[Custom],0)),"""")"
         End If
     End If
Next cell


End Sub

Sub clearContents_non_key()
    Dim ws As Worksheet
    Dim i As Long
    Sheets("Non-Key").Activate
    Set ws = ThisWorkbook.Worksheets("Non-Key") ' Replace "Sheet1" with the name of your worksheet
    
    For i = 7 To 43 ' Loop through
        If ws.Range("J" & i).value = "Uplift" Or ws.Range("J" & i).value = "D2C Conversion (Override)" Then
            ws.Range("K" & i & ":Y" & i).clearContents ' Clear the contents of the corresponding row in K2:Y2000
        End If
    Next i
End Sub

Sub insert_formulas_input()

 Dim lastCol As Long, i As Long
    Dim value  As Integer
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim rng As Range
    Dim cell As Range
    Dim myRange As Range
    
    Set ws = ThisWorkbook.Sheets("Koro")
   
    
lastRow = ws.Cells(ws.Rows.count, "J").End(xlUp).Row
Sheets("Koro").Activate
 
Set rng = ws.Range("J2:J" & lastRow)
Set myRange = ws.Range("K7:Y" & lastRow)
Set myRange = ws.Range("K7:Y" & lastRow)
For Each cell In rng
        If cell.value = "Sales Quantity Override" Then
             Range(Cells(cell.Row, 11), Cells(cell.Row, 25)).Formula = "=IFERROR(IF(INDEX(ip_live[0G_QABSMG],MATCH('Koro'!$H" & cell.Row & "&'Koro'!$I" & cell.Row & "&'Koro'!K$6,ip_live[index],0))=0,"""",INDEX(ip_live[0G_QABSMG],MATCH('Koro'!$H" & cell.Row & "&'Koro'!$I" & cell.Row & "&'Koro'!K$6,ip_live[index],0))),"""")"
        End If
Next cell
Sheets("Koro").Calculate
For i = myRange.Rows.count To 1 Step -1
        If Range("J" & i + 6).value = "Uplift" Or Range("J" & i + 6).value = "Paid Search % (Input)" Or Range("J" & i + 6).value = "Email % (Input)" Or Range("J" & i + 6).value = "Social % (Input)" Or Range("J" & i + 6).value = "D2C Conversion (Override)" Or Range("J" & i + 6).value = "Sales Quantity Override" Then
            myRange.Rows(i).value = myRange.Rows(i).value
        End If
         
Next i
End Sub

Sub insert_formulas_input_refresh()

 Dim lastCol As Long, i As Long
    Dim value  As Integer
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim rng As Range
    Dim cell As Range
    Dim myRange As Range
    
    Set ws = ThisWorkbook.Sheets("Koro")
   
    
lastRow = ws.Cells(ws.Rows.count, "J").End(xlUp).Row
Sheets("Koro").Activate
 
Set rng = ws.Range("J2:J" & lastRow)
Set myRange = ws.Range("K7:Y" & lastRow)
For Each cell In rng
        If cell.value = "Sales Quantity Override" Then
             Range(Cells(cell.Row, 11), Cells(cell.Row, 25)).Formula = "=IFERROR(IF(INDEX(ip_history[0G_QABSMG],MATCH('Koro'!$I" & cell.Row & "&'Koro'!K$6,ip_history[index],0))=0,"""",VLOOKUP('Koro'!$I" & cell.Row & "&'Koro'!K$6,ip_history!$F:$G,2,0)),"""")"
        End If
Next cell
Sheets("Koro").Calculate
For i = myRange.Rows.count To 1 Step -1
        If Range("J" & i + 6).value = "Uplift" Or Range("J" & i + 6).value = "Paid Search % (Input)" Or Range("J" & i + 6).value = "Email % (Input)" Or Range("J" & i + 6).value = "Social % (Input)" Or Range("J" & i + 6).value = "D2C Conversion (Override)" Or Range("J" & i + 6).value = "Sales Quantity Override" Then
            myRange.Rows(i).value = myRange.Rows(i).value
        End If
         
Next i
End Sub

Sub addcols_bm()                                        'modified by Noe, 31/03/2023

Application.Calculation = xlCalculationManual

Dim i As Long
i = 1

Do
i = i + 1
Loop Until bm_extract_query.Cells(1, i + 1) = ""

If bm_extract_query.Cells(1, i) = "Buffer5" Then
'do nothing
Else
 
    bm_extract_query.Cells(1, i + 1).value = "Buffer1"
    bm_extract_query.Cells(1, i + 2).value = "Buffer2"
    bm_extract_query.Cells(1, i + 3).value = "Buffer3"
    bm_extract_query.Cells(1, i + 4).value = "Buffer4"
    bm_extract_query.Cells(1, i + 5).value = "Buffer5"
    
    

Range("extract_basic_material_query[Buffer1]").FormulaR1C1 = "=IFERROR(R[-1]C+1,1)"

End If

Application.Calculation = xlCalculationAutomatic

End Sub


Sub addcols_mat()                                   'modified by Noe, 31/03/2023

Application.Calculation = xlCalculationManual

Dim i As Long
i = 1

Do
i = i + 1
Loop Until m_extract_query.Cells(1, i + 1) = ""

If m_extract_query.Cells(1, i) = "Buffer5" Then
'do nothing
Else
   
    m_extract_query.Cells(1, i + 1).value = "Buffer1"
    m_extract_query.Cells(1, i + 2).value = "Buffer2"
    m_extract_query.Cells(1, i + 3).value = "Buffer3"
    m_extract_query.Cells(1, i + 4).value = "Buffer4"
    m_extract_query.Cells(1, i + 5).value = "Buffer5"
    
Range("extract_material_query[Buffer1]").FormulaR1C1 = "=IFERROR(R[-1]C+1,1)"
   
End If

Application.Calculation = xlCalculationAutomatic

End Sub

