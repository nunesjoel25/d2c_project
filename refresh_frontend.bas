Attribute VB_Name = "refresh_frontend"
Sub refresh_key()

    ThisWorkbook.Unprotect "clausus"

    For Each wsheet In ActiveWorkbook.Sheets
    wsheet.Unprotect "clausus"
    Next wsheet

    Dim progress As Integer
    Dim regen_var As New frontendRegeneration
    Dim rng_koro As Range
    Dim rng_input As Range
    Dim rng_non_key As Range
      
    lastRow = Cells(Rows.count, "D").End(xlUp).Row

    Set rng_koro = KoroSheet.Range("F7:F" & lastRow)
    regen_var.unprotect_key_sheet "Koro"


    U_dialogbox.Show vbModeless
    progress = 10
    U_dialogbox.UpdateProgress progress, "Connecting to data source"
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayStatusBar = False
    ActiveWindow.FreezePanes = False
    
    If Sheets("Koro").AutoFilterMode Then Sheets("Koro").AutoFilter.ShowAllData
    Call frontend_regen.UnhideAllSheets
    Call frontend_regen.RefreshQuery

    progress = 60
    U_dialogbox.UpdateProgress progress, "Front end regeneration.."

    Call frontend_regen.clearAll
    Call frontend_regen.copyRange
    Call frontend_regen.MaterialListcopy
    Call frontend_regen.copy_tables
    Call formulas_to_values.ConvertRangeToValues_key_template
    
    progress = 75
    U_dialogbox.UpdateProgress progress, "Inserting Formulas.."
    
    Call insert_formulas.insert_formulas_input_refresh
    Call insert_formulas.InsertFormula
     
    progress = 80
    U_dialogbox.UpdateProgress progress, "Refreshing Total.."

    
    progress = 90
    U_dialogbox.UpdateProgress progress, "Finishing up.."


    regen_var.clearAll Range("NonKeyCellDefination"), NonKey, 0
    
    ThisWorkbook.Unprotect "clausus"
    
    regen_var.turn_very_hidden KoroSheet, TotalSheet, UserSelections

    If Not ActiveSheet.AutoFilterMode Then
        ActiveSheet.Range("KoroAutoFilterRange").AutoFilter
    End If
    
    regen_var.protect_sheet_key "Koro"
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.DisplayStatusBar = True
    
     frontend_regen.formatSheet KoroSheet
    progress = 100
    
    U_dialogbox.UpdateProgress progress, "Completed"
    
    Unload U_dialogbox

ThisWorkbook.Protect "clausus"
End Sub
Sub run_all_non_key()

Dim var As New frontendRegeneration
ThisWorkbook.Unprotect "clausus"
var.unprotect_key_sheet "Non-Key"
Dim progress As Integer
Dim regen_var As New frontendRegeneration

    U_dialogbox.Show vbModeless
    progress = 10
    U_dialogbox.UpdateProgress progress, "Connecting to data source"
    
    For Each wsheet In ActiveWorkbook.Sheets
    wsheet.Unprotect "clausus"
    Next wsheet

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayStatusBar = False
    ActiveWindow.FreezePanes = False

If NonKey.AutoFilterMode Then NonKey.AutoFilter.ShowAllData

Call frontend_regen.UnhideAllSheets
    progress = 25
Call frontend_regen.RefreshQuery
    progress = 70
    U_dialogbox.UpdateProgress progress, "Front end regeneration.."

Call frontend_regen.clearAll
Call frontend_regen.copyRange
Call frontend_regen.MaterialListcopy
Call frontend_regen.copy_tables
Call formulas_to_values.ConvertRangeToValues_non_key_template

    progress = 75
    
    U_dialogbox.UpdateProgress progress, "Refreshing Total.."
    
Call refresh_total

ThisWorkbook.Unprotect "clausus"
    progress = 80
    U_dialogbox.UpdateProgress progress, "Tyding up.."

Application.Calculation = xlCalculationAutomatic

Sheets("Non-Key").Activate

regen_var.turn_very_hidden NonKey, TotalSheet, UserSelections
Application.ScreenUpdating = True
Application.DisplayStatusBar = True
    progress = 100
    U_dialogbox.UpdateProgress progress, "Completed"
    
    Unload U_dialogbox
     
    If Not ActiveSheet.AutoFilterMode Then
    
    ActiveSheet.Range("C6:AD6").AutoFilter
    
    End If
    frontend_regen.formatSheet NonKey
    regen_var.protect_sheet_key "Non-Key"
    
Sheets("Non-Key").Calculate

ThisWorkbook.Protect "clausus"

End Sub

Sub refresh_all()
Dim clear As New frontendRegeneration
     Select Case clear.UserSelectionLevel
     Case Is = "Sales Organisation"
        refresh_input
    Case Is = "SeAG"
        Select Case clear.UserSelectionIndicator
            Case Is = "Key"
                refresh_key
            Case Is = "Non-Key"
                run_all_non_key
        End Select
    End Select
 
End Sub
Sub refresh_total()
Dim var As New frontendRegeneration
ThisWorkbook.Unprotect "clausus"
var.unprotect_key_sheet "Total"
Application.Calculation = xlCalculationManual
If Sheets("Total").AutoFilterMode Then Sheets("Total").AutoFilter.ShowAllData
Call total_regen.Unhide_Columns_Containing_Value_total
Call total_regen.clear_total
Call total_regen.copyRange_total
Call total_regen.MaterialListcopy_total
Call total_regen.HighlightTextWithAsterisks_total
Call total_regen.Hide_Columns_Containing_Value_total
Call total_regen.calculate_pivot_sheet
Call total_regen.refresh_Pivot
Call formulas_to_values.ConvertRangeToValues_total
Call total_regen.GroupRowsWithAsterisk_total
Sheets("Total").Activate
If Not ActiveSheet.AutoFilterMode Then
    ActiveSheet.Range("J5:X5").AutoFilter
  End If
Application.Calculation = xlCalculationAutomatic
var.protect_sheet_key "Total"
ThisWorkbook.Protect "clausus"
End Sub

Sub on_demand_refresh()
If Worksheets("User Selections").Range("G7").value = "Key" Then
   refresh_key_on_demand
   
   Else
   run_all_non_key_on_demand
   End If
End Sub

Sub clear_filters()
Dim var As New frontendRegeneration
ThisWorkbook.Unprotect "clausus"
If Worksheets("User Selections").Range("G7").value = "Key" Then
var.unprotect_key_sheet "Koro"
If Sheets("Koro").AutoFilterMode Then Sheets("Koro").AutoFilter.ShowAllData
var.protect_sheet_key "Koro"

Else
var.unprotect_key_sheet "Non-Key"
If Sheets("Non-Key").AutoFilterMode Then Sheets("Non-Key").AutoFilter.ShowAllData
var.protect_sheet_key "Non-Key"
End If
ThisWorkbook.Protect "clausus"
End Sub

Sub training_link()
Dim URL As String
    URL = "https://sonyeur.sharepoint.com/sites/S022-SupplyChain/SCMPlatform/Shared%20Documents/Forms/AllItems.aspx?csf=1&web=1&e=AnRrfn&cid=02813ce6%2Dfea0%2D4027%2D8606%2D01f3a9a2a5f9&FolderCTID=0x01200024286FA772E8524BAD559FE9BB5248BE&id=%2Fsites%2FS022%2DSupplyChain%2FSCMPlatform%2FShared%20Documents%2FSeihan%20Standardisation%2F60%5FBusiness%20Operation%20Files%2F00%5FCommon%2FD2C%20Marketing%20Template%2FDocumentation&viewid=ce539b45%2D97e2%2D4514%2Dbb50%2D473793abd40d"

    ActiveWorkbook.FollowHyperlink URL

End Sub
Sub refresh_input()
    
    ThisWorkbook.Unprotect "clausus"
    Dim progress As Integer
    Dim regen_var As New frontendRegeneration
    regen_var.unprotect_key_sheet "Input Sheet"

    
    U_dialogbox.Show vbModeless
    progress = 10
    U_dialogbox.UpdateProgress progress, "Connecting to data source"
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayStatusBar = False
    ActiveWindow.FreezePanes = False
    
    If Sheets("Input Sheet").AutoFilterMode Then Sheets("Input Sheet").AutoFilter.ShowAllData
    
    Call frontend_regen.UnhideAllSheets
    Call frontend_regen.RefreshQuery

    progress = 60
    U_dialogbox.UpdateProgress progress, "Front end regeneration.."
    
    For Each wsheet In ActiveWorkbook.Sheets
    wsheet.Unprotect "clausus"
    Next wsheet

    Call frontend_regen.clearAll
    Call frontend_regen.copyRange
    Call frontend_regen.MaterialListcopy
    Call frontend_regen.copy_tables
    Call frontend_regen.GroupColumns_input_sheet
    Call formulas_to_values.ConvertRangeToValues_input_Sheet
    
    progress = 75
    U_dialogbox.UpdateProgress progress, "Finishing up.."


    regen_var.clearAll Range("NonKeyCellDefination"), NonKey, 0
    regen_var.clearAll Range("input_grid_key"), KoroSheet, 1
    
    ThisWorkbook.Unprotect "clausus"
    
    regen_var.turn_very_hidden InputSheet, UserSelections

    If Not InputSheet.AutoFilterMode Then
        InputSheet.Range("InputSheetAutoFilterRange").AutoFilter
    End If
     frontend_regen.formatSheet InputSheet
    regen_var.protect_sheet_key "Input Sheet"
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.DisplayStatusBar = True
       
    progress = 100
    
    U_dialogbox.UpdateProgress progress, "Completed"
    
    Unload U_dialogbox

ThisWorkbook.Protect "clausus"
End Sub

Sub run_all_non_key_on_demand()
Dim regen_var As New frontendRegeneration

regen_var.unprotect_key_sheet "Non-Key"

ThisWorkbook.Unprotect "clausus"

Dim progress As Integer
    U_dialogbox.Show vbModeless
    progress = 10
    U_dialogbox.UpdateProgress progress, "Connecting to data source"

Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.DisplayStatusBar = False


Call frontend_regen.UnhideAllSheets
    progress = 25
    U_dialogbox.UpdateProgress progress, "Refreshing sub data queries.."
 regen_var.QueryRefresh 1
 
    progress = 30
    U_dialogbox.UpdateProgress progress, "Refreshing Orders and Traffic data query.."
 regen_var.QueryRefresh 3
    progress = 45
    U_dialogbox.UpdateProgress progress, "Refreshing MA Calcs...."
 regen_var.QueryRefresh 11
    

    Call frontend_regen.copy_tables
   
    progress = 70
    U_dialogbox.UpdateProgress progress, "Refreshing Total.."
    Call refresh_total
ThisWorkbook.Unprotect "clausus"
    progress = 80
    U_dialogbox.UpdateProgress progress, "Tyding up.."
    
 regen_var.turn_very_hidden NonKey, TotalSheet, UserSelections


    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.DisplayStatusBar = True
    progress = 100
    
    U_dialogbox.UpdateProgress progress, "Completed"
    UploadConfig.Visible = xlSheetVeryHidden
    Unload U_dialogbox
    
    regen_var.protect_sheet_key "Non-Key"
    
    ThisWorkbook.Protect "clausus"
End Sub
Sub refresh_key_on_demand()
    
    Dim regen_var As New frontendRegeneration

    regen_var.unprotect_key_sheet "Koro"

    ThisWorkbook.Unprotect "clausus"

    Dim progress As Integer
    U_dialogbox.Show vbModeless
    progress = 10
    U_dialogbox.UpdateProgress progress, "Connecting to data source"

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayStatusBar = False


    Call frontend_regen.UnhideAllSheets
    progress = 25
    U_dialogbox.UpdateProgress progress, "Refreshing sub data queries.."
    regen_var.QueryRefresh 1
 
    progress = 30
    U_dialogbox.UpdateProgress progress, "Refreshing Orders and Traffic data query.."
    regen_var.QueryRefresh 3
    progress = 45
    
    Call frontend_regen.copy_tables
    Call insert_formulas.InsertFormula
   
    progress = 70
    U_dialogbox.UpdateProgress progress, "Refreshing Total.."
    Call refresh_total
    
    ThisWorkbook.Unprotect "clausus"
    progress = 80
    U_dialogbox.UpdateProgress progress, "Tyding up.."
    
    regen_var.turn_very_hidden KoroSheet, TotalSheet, UserSelections


    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.DisplayStatusBar = True
    progress = 100
    
    U_dialogbox.UpdateProgress progress, "Completed"
    UploadConfig.Visible = xlSheetVeryHidden
    Unload U_dialogbox
    
    regen_var.protect_sheet_key "Koro"
    
    ThisWorkbook.Protect "clausus"
End Sub


