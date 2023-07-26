Attribute VB_Name = "extract_total"
Sub extract_total_key()
    Dim regen_var As New frontendRegeneration
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim clearRange As Range
    Dim currentDate As Date
    Dim cl As Range
    Dim copyRange As Range

    ThisWorkbook.Unprotect "clausus"
    regen_var.unprotect_key_sheet "Koro"

    KoroSheet.Activate
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    KoroSheet.Range("K11").Select
    ActiveWindow.FreezePanes = False
    bm_rw_extract.Visible = True
    KoroSheet.Cells.ClearOutline
    
    bm_rw_extract.Cells.clear
    m_rw_extract.Cells.clearContents
    'bm_rw_extract.ListObjects("extract_basic_material_raw_data").Delete
    'm_rw_extract.ListObjects("extract_material_raw_data").Delete
    regen_var.u_hide_columns_rows KoroSheet, False
 
    ' Set the worksheet
    Set ws = ThisWorkbook.Worksheets("Koro")
    Set tgt_bm = ThisWorkbook.Worksheets("Raw_data_extract_basic_material")
    Set tgt_m = ThisWorkbook.Worksheets("Raw_data_extract_material")
'
    ' Set the last row of data
    lastRow = ws.Cells(ws.Rows.count, "J").End(xlUp).Row

    ws.Range("C6:AD" & lastRow).AutoFilter field:=4, Criteria1:="."
    Set copyRange = ws.Range("C6:AD" & lastRow)

    copyRange.SpecialCells(xlCellTypeVisible).Copy
    tgt_bm.Cells(1, 1).PasteSpecial Paste:=xlPasteValues
    tgt_bm.Columns("D:G").EntireColumn.Delete
    tgt_bm.Columns("T:X").EntireColumn.Delete
    
    ws.AutoFilterMode = False
    ws.Range("C6:AD" & lastRow).AutoFilter field:=4, Criteria1:=".."

    copyRange.SpecialCells(xlCellTypeVisible).Copy
    tgt_m.Cells(1, 1).PasteSpecial Paste:=xlPasteValues
    tgt_m.Columns("B:F").EntireColumn.Delete
    tgt_m.Columns("S:W").EntireColumn.Delete

    

    ws.AutoFilterMode = False
    frontend_regen.formatSheet KoroSheet
    ws.Range("C6:AD" & lastRow).AutoFilter

    tgt_m.ListObjects.Add(xlSrcRange, Range("raw_extract_material"), , xlYes).Name = "extract_material_raw_data"
    tgt_bm.ListObjects.Add(xlSrcRange, Range("raw_extract"), , xlYes).Name = "extract_basic_material_raw_data"
    
    m_extract_query.ListObjects("extract_material_query") _
    .QueryTable.Refresh BackgroundQuery:=False
    bm_extract_query.ListObjects("extract_basic_material_query") _
    .QueryTable.Refresh BackgroundQuery:=False
 
 regen_var.protect_sheet_key "Koro"
End Sub

Sub extract_total_nonkey()
 Dim regen_var As New frontendRegeneration
 Dim ws As Worksheet
 Dim lastRow As Long
 Dim clearRange As Range
 Dim currentDate As Date
 Dim cl As Range
 Dim copyRange As Range

    ThisWorkbook.Unprotect "clausus"
    regen_var.unprotect_key_sheet "Non-Key"
 
    Worksheets("Non-Key").Activate
    Application.ScreenUpdating = False

    Sheets("Non-Key").Range("K7").Select
    ActiveWindow.FreezePanes = False
    Application.Calculation = xlCalculationManual
    bm_rw_extract.Visible = True
    NonKey.Cells.ClearOutline
    bm_rw_extract.Cells.clear
    
    bm_rw_extract.Cells.clear
    m_rw_extract.Cells.clearContents
    'bm_rw_extract.ListObjects("extract_basic_material_raw_data").Delete
    'm_rw_extract.ListObjects("extract_material_raw_data").Delete
    regen_var.u_hide_columns_rows NonKey, False
 
    ' Set the worksheet
    Set ws = ThisWorkbook.Worksheets("Non-Key")
    Set tgt_bm = bm_rw_extract
    Set tgt_m = m_rw_extract
'
    ' Set the last row of data
    lastRow = ws.Cells(ws.Rows.count, "J").End(xlUp).Row

    ws.Range("C6:AD" & lastRow).AutoFilter field:=4, Criteria1:="."
    Set copyRange = ws.Range("C6:AD" & lastRow)
    

    copyRange.SpecialCells(xlCellTypeVisible).Copy
    tgt_bm.Cells(1, 1).PasteSpecial Paste:=xlPasteValues
    tgt_bm.Columns("D:G").EntireColumn.Delete
    tgt_bm.Columns("T:X").EntireColumn.Delete
    
    ws.Range("A6:AD" & lastRow).AutoFilter field:=4, Criteria1:=".."
    copyRange.SpecialCells(xlCellTypeVisible).Copy
    tgt_m.Cells(1, 1).PasteSpecial Paste:=xlPasteValues
    tgt_m.Columns("B:F").EntireColumn.Delete
    tgt_m.Columns("S:W").EntireColumn.Delete
    

    ws.AutoFilterMode = False
    frontend_regen.formatSheet NonKey
    ws.Range("C6:AD" & lastRow).AutoFilter

    tgt_bm.ListObjects.Add(xlSrcRange, Range("raw_extract"), , xlYes).Name = "extract_basic_material_raw_data"
    Worksheets("extract_basic_material").ListObjects("extract_basic_material_query") _
    .QueryTable.Refresh BackgroundQuery:=False
    
    tgt_m.ListObjects.Add(xlSrcRange, Range("raw_extract_material"), , xlYes).Name = "extract_material_raw_data"
    
    Worksheets("extract_material").ListObjects("extract_material_query") _
    .QueryTable.Refresh BackgroundQuery:=False

 regen_var.protect_sheet_key "Non-key"
End Sub

Sub extract_total_key_material()
    Dim regen_var As New frontendRegeneration
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim clearRange As Range
    Dim currentDate As Date
    Dim cl As Range
    Dim copyRange As Range

    ThisWorkbook.Unprotect "clausus"
    regen_var.unprotect_key_sheet "Koro"

    KoroSheet.Activate
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    KoroSheet.Range("K11").Select
    ActiveWindow.FreezePanes = False
    m_rw_extract.Visible = True
    KoroSheet.Cells.ClearOutline
    'm_rw_extract.ListObjects("Raw_data_extract_material").Delete

    regen_var.u_hide_columns_rows KoroSheet, False
    Sheets("Raw_data_extract_material").Range("A1:AC10000").clear
 
    Set ws = ThisWorkbook.Worksheets("Koro")
    Set tgt = ThisWorkbook.Worksheets("Raw_data_extract_material")


    lastRow = ws.Cells(ws.Rows.count, "J").End(xlUp).Row

    ws.Range("A6:AD" & lastRow).AutoFilter field:=4, Criteria1:=".."

    Set copyRange = ws.Range("A6:AD" & lastRow)

    copyRange.SpecialCells(xlCellTypeVisible).Copy
    tgt.Cells(1, 1).PasteSpecial Paste:=xlPasteValues
    tgt.Columns("D:H").EntireColumn.Delete
    tgt.Columns("U:Y").EntireColumn.Delete


    frontend_regen.formatSheet KoroSheet
    ws.Range("C6:AD" & lastRow).AutoFilter

    tgt.ListObjects.Add(xlSrcRange, Range("raw_extract_material"), , xlYes).Name = "extract_material_raw_data"
    
    Worksheets("extract_material").ListObjects("extract_material_query") _
    .QueryTable.Refresh BackgroundQuery:=False

    Worksheets("Raw_data_extract_basic_material").Visible = xlSheetVeryHidden

    regen_var.protect_sheet_key "Koro"
End Sub

Sub extract_total_non_key_material()
 Dim regen_var As New frontendRegeneration
 Dim ws As Worksheet
 Dim lastRow As Long
 Dim clearRange As Range
 Dim currentDate As Date
 Dim cl As Range
 Dim copyRange As Range

    ThisWorkbook.Unprotect "clausus"
    regen_var.unprotect_key_sheet "Non-Key"
 
    Worksheets("Non-Key").Activate
    Application.ScreenUpdating = False

    Sheets("Non-Key").Range("K7").Select
    ActiveWindow.FreezePanes = False
    Application.Calculation = xlCalculationManual
    m_rw_extract.Visible = True
    NonKey.Cells.ClearOutline
    bm_rw_extract.Cells.clear
    
    bm_rw_extract.Cells.clear
    m_rw_extract.Cells.clearContents
    'bm_rw_extract.ListObjects("extract_basic_material_raw_data").Delete
    'm_rw_extract.ListObjects("extract_material_raw_data").Delete
    regen_var.u_hide_columns_rows NonKey, False
 
    ' Set the worksheet
    Set ws = ThisWorkbook.Worksheets("Non-Key")
    Set tgt_bm = bm_rw_extract
    Set tgt_m = m_rw_extract
'
    ' Set the last row of data
    lastRow = ws.Cells(ws.Rows.count, "J").End(xlUp).Row

    ws.Range("C6:AD" & lastRow).AutoFilter field:=4, Criteria1:=".."
    Set copyRange = ws.Range("C6:AD" & lastRow)

    copyRange.SpecialCells(xlCellTypeVisible).Copy
    tgt_m.Cells(1, 1).PasteSpecial Paste:=xlPasteValues
    tgt_m.Columns("B:F").EntireColumn.Delete
    tgt_m.Columns("T:X").EntireColumn.Delete

    frontend_regen.formatSheet NonKey
    ws.Range("C6:AD" & lastRow).AutoFilter

    tgt_m.ListObjects.Add(xlSrcRange, Range("raw_extract_material"), , xlYes).Name = "extract_material_raw_data"
    
    m_extract_query.ListObjects("extract_material_query") _
    .QueryTable.Refresh BackgroundQuery:=False


    regen_var.protect_sheet_key "Non-Key"
End Sub

Sub extract_total_nonkey_retrival()
 unprotect_key_sheet "Non-key"
 Worksheets("Non-Key").Activate
 Application.ScreenUpdating = False
Sheets("Non-Key").Range("K7").Select
ActiveWindow.FreezePanes = False
Application.Calculation = xlCalculationManual
Worksheets("Raw_data_extract_ret").Visible = True
Worksheets("Non-Key").Cells.ClearOutline
Worksheets("Raw_data_extract_ret").ListObjects("extract_ret").Delete
Call nonkey_refresh.Unhide_Columns_Containing_Value_non_key
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim clearRange As Range
    Dim currentDate As Date
    Dim cl As Range
    Dim copyRange As Range

    ' Set the worksheet
    Set ws = ThisWorkbook.Worksheets("Non-Key")
    Set tgt = ThisWorkbook.Worksheets("Raw_data_extract_ret")
'
    ' Set the last row of data
    lastRow = ws.Cells(ws.Rows.count, "J").End(xlUp).Row

    ws.Range("A6:AD" & lastRow).AutoFilter field:=4, Criteria1:="."
Set copyRange = ws.Range("A6:AD" & lastRow)

copyRange.SpecialCells(xlCellTypeVisible).Copy
tgt.Cells(1, 1).PasteSpecial Paste:=xlPasteValues
tgt.Columns("F:I").EntireColumn.Delete
tgt.Columns("V:Z").EntireColumn.Delete

ws.AutoFilterMode = False
Call nonkey_refresh.GroupRowsWithAsterisk_non_key
Call nonkey_refresh.Hide_Columns_Containing_Value_non_key
ws.Range("C6:AD" & lastRow).AutoFilter

tgt.ListObjects.Add(xlSrcRange, Range("raw_extract_ret"), , xlYes).Name = "extract_ret"
Worksheets("extract_ret").ListObjects("extract_ret_2").QueryTable.Refresh BackgroundQuery:=False

Worksheets("Raw_data_extract_ret").Visible = xlSheetVeryHidden
Call freeze_panes.FreezePanes_non_key
protect_sheet_key "Non-key"
End Sub

Sub extract_input_sheet()
    Dim regen_var As New frontendRegeneration
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim clearRange As Range
    Dim currentDate As Date
    Dim cl As Range
    Dim copyRange As Range

    'ThisWorkbook.Unprotect "clausus"
    regen_var.unprotect_key_sheet "Input Sheet"
    ThisWorkbook.Unprotect "clausus"

    InputSheet.Activate
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    InputSheet.Range("M8").Select
    ActiveWindow.FreezePanes = False
    input_rw_extract.Visible = True
    raw_extract_input_query.Visible = True
    InputSheet.Cells.ClearOutline
    input_rw_extract.Cells.clear
    regen_var.u_hide_columns_rows InputSheet, False
 
   
    Set ws = ThisWorkbook.Worksheets("Input Sheet")
    Set tgt_input = ThisWorkbook.Worksheets("Input_Extract")

    lastRow = ws.Cells(ws.Rows.count, "H").End(xlUp).Row

    ws.Range("C7:AT" & lastRow).AutoFilter field:=7, Criteria1:="."
    Set copyRange = ws.Range("C7:AT" & lastRow)

    copyRange.SpecialCells(xlCellTypeVisible).Copy
    tgt_input.Cells(1, 1).PasteSpecial Paste:=xlPasteValues
    tgt_input.Columns("B:H").EntireColumn.Delete
    tgt_input.Columns("P:T").EntireColumn.Delete
    tgt_input.Columns("AB:AF").EntireColumn.Delete
    
    ws.AutoFilterMode = False
    regen_var.groupColumns
    frontend_regen.formatSheet InputSheet
    ws.Range("C7:AT" & lastRow).AutoFilter

    tgt_input.ListObjects.Add(xlSrcRange, Range("raw_extract_input"), , xlYes).Name = "raw_extract_input"

    
    raw_extract_input_query.ListObjects("raw_extract_input_query") _
    .QueryTable.Refresh BackgroundQuery:=False
 
 
 regen_var.protect_sheet_key "Input Sheet"
End Sub



