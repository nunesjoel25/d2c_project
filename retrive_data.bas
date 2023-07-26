Attribute VB_Name = "retrive_data"
Sub retrive_formulas_key()
ThisWorkbook.Unprotect "clausus"
Dim progress As Integer
unprotect_key_sheet "Koro"
   
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.DisplayStatusBar = False
ActiveWindow.FreezePanes = False
ThisWorkbook.Unprotect "clausus"
Worksheets("Koro_live").Visible = xlSheetVisible

Call refresh_queries.refresh_data_queries_retival
Call insert_formulas.insert_formulas_retrival
Call formulas_to_values.ConvertRangeToValues_key_template_retrival
Call insert_formulas.InsertFormula
Call freeze_panes.FreezePanes_key
Call frontend_regen.GroupRowsWithAsterisk
Call frontend_regen.Hide_Columns_Containing_Value

ActiveWindow.FreezePanes = True
Application.ScreenUpdating = True
Application.DisplayStatusBar = True
protect_sheet_key "Koro"
Application.Calculation = xlCalculationAutomatic
Sheets("Koro").Calculate
Worksheets("Koro_live").Visible = xlSheetVeryHidden
ThisWorkbook.Protect "clausus"
End Sub
Sub retrive_formulas_non_key()
ThisWorkbook.Unprotect "clausus"
Dim progress As Integer
unprotect_key_sheet "Non-Key"
   
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.DisplayStatusBar = False
ActiveWindow.FreezePanes = False
ThisWorkbook.Unprotect "clausus"
Worksheets("Koro_live").Visible = xlSheetVisible
Sheet36.Visible = xlSheetVisible
Sheet40.Visible = xlSheetVisible


Call refresh_queries.refresh_data_queries_retival
Call insert_formulas.insert_formulas_retrival_non_key
Call extract_total.extract_total_nonkey_retrival
Call refresh_queries.refresh_data_queries_retival_non_key
Call formulas_to_values.ConvertRangeToValues_non_key_template_retrival
Call freeze_panes.FreezePanes_non_key
Call nonkey_refresh.GroupRowsWithAsterisk_non_key
Call nonkey_refresh.HighlightTextWithAsterisks_non_key

ActiveWindow.FreezePanes = True
Application.ScreenUpdating = True
Application.DisplayStatusBar = True
protect_sheet_key "Non-Key"
Application.Calculation = xlCalculationAutomatic
Sheets("Non-Key").Calculate
Worksheets("Koro_live").Visible = xlSheetVeryHidden
Sheet36.Visible = xlSheetVeryHidden
Sheet40.Visible = xlSheetVeryHidden

ThisWorkbook.Protect "clausus"
End Sub

Sub retrive_data()

If Worksheets("User Selections").Range("G7").value = "Key" Then

retrive_formulas_key

Else
retrive_formulas_non_key
End If
End Sub

