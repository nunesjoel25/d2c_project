Attribute VB_Name = "formulas_to_values"
Sub ConvertRangeToValues_key_template()
 Dim ws As Worksheet
 Dim lastRow As Long
    ' Set the worksheet
    Set ws = ThisWorkbook.Worksheets("Koro")
  
    ' Set the last row of data
    lastRow = ws.Cells(ws.Rows.count, "J").End(xlUp).Row
    Dim other_range As Range
    Dim myRange As Range
    Set myRange = ws.Range("K6:AD" & lastRow)
    Set other_range = ws.Range("H2:I" & lastRow)
    Sheets("Koro").Calculate

    ' Loop through each row in the range and convert to values if the condition is met
    Dim i As Long
    For i = myRange.Rows.count To 1 Step -1
        If Range("H" & i + 5).value = "PY" Or Range("H" & i + 5).value = "OB" Or Range("H" & i + 5).value = Worksheets("User Selections").Range("E7").value Then
            myRange.Rows(i).value = myRange.Rows(i).value
        End If
    Next i
    
    For i = other_range.Rows.count To 1 Step -1
        other_range.Rows(i).value = other_range.Rows(i).value
    Next i
    
    For i = myRange.Rows.count To 1 Step -1
        If Range("J" & i + 5).value = "Sales Qty %" Or Range("J" & i + 5).value = "Last Reported Stock Qty" Then
            myRange.Rows(i).value = myRange.Rows(i).value
        End If
    Next i

End Sub
Sub ConvertRangeToValues_non_key_template()
 Dim ws As Worksheet
 Dim lastRow As Long
    ' Set the worksheet
    Set ws = ThisWorkbook.Worksheets("Non-Key")
  
    ' Set the last row of data
    lastRow = ws.Cells(ws.Rows.count, "J").End(xlUp).Row
    Dim other_range As Range
    Dim myRange As Range
    Set myRange = ws.Range("K43:AD" & lastRow)
    Set other_range = ws.Range("H43:I" & lastRow)
    Sheets("Non-Key").Calculate

    ' Loop through each row in the range and convert to values if the condition is met
    Dim i As Long
    For i = myRange.Rows.count To 1 Step -1
        If Range("H" & i + 5).value = "OB" Or Range("H" & i + 5).value = Worksheets("User Selections").Range("E7").value Then
            myRange.Rows(i).value = myRange.Rows(i).value
        End If
    Next i
    
    For i = other_range.Rows.count To 1 Step -1
        other_range.Rows(i).value = other_range.Rows(i).value
    Next i

End Sub
Sub ConvertRangeToValues_total()
 Dim ws As Worksheet
 Dim lastRow As Long

    ' Set the worksheet
    Set ws = ThisWorkbook.Worksheets("Total")
  
    ' Set the last row of data
    lastRow = ws.Cells(ws.Rows.count, "I").End(xlUp).Row
    Dim other_range As Range
    Dim myRange As Range
    Set myRange = ws.Range("H6:AC" & lastRow)
    Sheets("Total").Calculate

    ' Loop through each row in the range and convert to values if the condition is met
    Dim i As Long
    For i = myRange.Rows.count To 1 Step -1
            myRange.Rows(i).value = myRange.Rows(i).value
    Next i
    

End Sub
Sub ConvertRangeToValues_key_template_retrival()
 Dim ws As Worksheet
 Dim lastRow As Long
  
    ' Set the worksheet
    Set ws = ThisWorkbook.Worksheets("Koro")
  
    ' Set the last row of data
    lastRow = ws.Cells(ws.Rows.count, "J").End(xlUp).Row
    Dim other_range As Range
    Dim myRange As Range
    Set myRange = ws.Range("K7:Y" & lastRow)
    Sheets("Koro").Calculate

    ' Loop through each row in the range and convert to values if the condition is met
    Dim i As Long
    For i = myRange.Rows.count To 1 Step -1
        If Range("J" & i + 6).value = "Uplift" Or Range("J" & i + 6).value = "Paid Search % (Input)" Or Range("J" & i + 6).value = "Email % (Input)" Or Range("J" & i + 6).value = "Social % (Input)" Or Range("J" & i + 6).value = "D2C Conversion (Override)" Or Range("J" & i + 6).value = "Sales Quantity Override" Then
            myRange.Rows(i).value = myRange.Rows(i).value
        End If
    Next i

End Sub

Sub ConvertRangeToValues_non_key_template_retrival()
 Dim ws As Worksheet
 Dim lastRow As Long
   
    ' Set the worksheet
    Set ws = ThisWorkbook.Worksheets("Non-Key")
  
    ' Set the last row of data
    lastRow = ws.Cells(ws.Rows.count, "J").End(xlUp).Row
    Dim other_range As Range
    Dim myRange As Range
    Set myRange = ws.Range("K7:Y" & lastRow)
    Sheets("Non-Key").Calculate

    ' Loop through each row in the range and convert to values if the condition is met
    Dim i As Long
    For i = myRange.Rows.count To 1 Step -1
        If Range("J" & i + 6).value = "Uplift" Or Range("J" & i + 6).value = "D2C Conversion (Override)" Then
            myRange.Rows(i).value = myRange.Rows(i).value
        End If
    Next i

End Sub

Sub ConvertRangeToValues_input_Sheet()
 Dim ws As Worksheet
 Dim lastRow As Long
 
    ' Set the worksheet
    Set ws = ThisWorkbook.Worksheets("Input Sheet")
  
    ' Set the last row of data
    lastRow = ws.Cells(ws.Rows.count, "D").End(xlUp).Row
    Dim myRange As Range
    Set myRange = ws.Range("K8:AT" & lastRow)
    Sheets("Input Sheet").Calculate

    ' Loop through each row in the range and convert to values if the condition is met
    Dim i As Long
     
    For i = myRange.Rows.count To 1 Step -1
        If Range("L" & i + 7).value = "Sell in Quantity Override" Or Range("L" & i + 7).value = "SAP Inventory" Or Range("L" & i + 7).value = "Last Reported Stock(Hybris)" Or Range("L" & i + 7).value = "Actual Replenishment Qty " Or Range("L" & i + 7).value = "Actual Sell In Qty" Then
            myRange.Rows(i).value = myRange.Rows(i).value
        End If
    Next i

End Sub


