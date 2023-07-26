Attribute VB_Name = "ip_extract"
Sub ip_extract_call()
    Dim progress As Integer
    
    U_dialogbox.Show vbModeless
    progress = 10
    U_dialogbox.UpdateProgress progress, "Connecting to data source"
    ThisWorkbook.Unprotect "clausus"

    If ActiveSheet.AutoFilterMode Then ActiveSheet.AutoFilter.ShowAllData

    progress = 45
    
    U_dialogbox.UpdateProgress progress, "Extracting data.."
    Call extract_basic_material

    ActiveWorkbook.Worksheets("fromDate").ListObjects("fromDate") _
    .QueryTable.Refresh BackgroundQuery:=False
    ActiveWorkbook.Worksheets("toDate").ListObjects("toDate") _
    .QueryTable.Refresh BackgroundQuery:=False
    progress = 75
    
    U_dialogbox.UpdateProgress progress, "Writing file to txt..."
    
    Call write_to_txt
    
    ThisWorkbook.Protect "clausus"
    MsgBox "File extract complete" & vbNewLine & "Note: Extract is saved in the local directory as the Workbook.", vbInformation, "Upload Successful"
    progress = 100
    
    U_dialogbox.UpdateProgress progress, "Completed"
    Unload U_dialogbox

End Sub
Sub extract_basic_material()
    Dim var As New frontendRegeneration
    Dim progress As Integer
    
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
   
    U_dialogbox.Show vbModeless
    progress = 10
    U_dialogbox.UpdateProgress progress, "Connecting to data source"
 
    ThisWorkbook.Unprotect "clausus"
    If Sheets("Koro").AutoFilterMode Then Sheets("Koro").AutoFilter.ShowAllData
    If Sheets("Non-Key").AutoFilterMode Then Sheets("Non-Key").AutoFilter.ShowAllData
    
    frontend_regen.UnhideAllSheets

    If Worksheets("User Selections").Range("G7").value = "Key" Then
       progress = 30
    U_dialogbox.UpdateProgress progress, "Extracting data"
        
        Call extract_total.extract_total_key: Call addcols_bm: Call addcols_mat      'modified by Noe, 31/03/2023
         
         var.turn_very_hidden KoroSheet, UserSelections, TotalSheet
         var.u_hide_columns_rows KoroSheet, True
    Else
    
        Call extract_total.extract_total_nonkey: Call addcols_bm: Call addcols_mat
         
         var.turn_very_hidden NonKey, UserSelections, TotalSheet
         var.u_hide_columns_rows NonKey, True
        
    End If
 
  progress = 60
    
    U_dialogbox.UpdateProgress progress, "Uploading data"
 
    Call UploadData                                                    'modified by Noe, 31/03/2023
 
    ThisWorkbook.Protect "clausus"
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    progress = 100
    
    U_dialogbox.UpdateProgress progress, "Completed"
    Unload U_dialogbox

End Sub

Sub write_to_txt()

Dim myrng As Range

Dim i, j
  Dim datim As String
    datim = Format(CStr(Now), "yyyy_mm_dd_hh_mm_ss")

Filename = ThisWorkbook.Path & "\" & "IP extract" & datim & ".txt"

Open Filename For Output As #1

Set myrng = Range("extract_material_query")

   

For i = 1 To myrng.Rows.count + 1

'For j = 1 To myrng.Columns.count   'modified by Noe, 31/03/2023
For j = 1 To 7                      'modified by Noe, 31/03/2023

lineText = IIf(j = 1, "", lineText & Chr(9)) & myrng.Cells(i - 1, j)

            Next j

Print #1, lineText

            If i > 2 And Cells(i, 3).value <> Cells(i + 1, 3) Then


Print #1, lineText

            Else

                  'Do Nothing

            End If

Next i

      Close #1
End Sub

Sub UploadToServer()
   
    Dim response As VbMsgBoxResult
    

    response = MsgBox("Are you sure you want to upload to server?", vbQuestion + vbYesNo, "Confirmation")

    If response = vbYes Then
        If Worksheets("User Selections").Range("C6").value = "SeAG" Then
    
        Call extract_basic_material
        
        Else
        
        Call input_sheet_upload
        MsgBox "File uploaded to server successfully!", vbInformation, "Upload Successful"
        
        End If
    Else
       
        MsgBox "File was not uploaded to server.", vbExclamation, "Upload Canceled"
    End If
End Sub
Sub refreshfromServer()
 
    Dim response As VbMsgBoxResult
    
    response = MsgBox("Are you sure you want to Refresh from Server?" & vbNewLine & "Note: By doing so you will loose your latest planning data!.", vbQuestion + vbYesNo, "Confirmation")

    If response = vbYes Then
        Call refresh_frontend.refresh_all
        MsgBox "File refresh from server successfully!", vbInformation, "Refresh Successful"
    Else
       
        MsgBox "File was not refreshed, procedure was cancelled", vbExclamation, "Refresh Canceled"
    End If
End Sub

Sub input_sheet_upload()
    Dim var As New frontendRegeneration
    Dim progress As Integer
    
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
   
    U_dialogbox.Show vbModeless
    progress = 10
    U_dialogbox.UpdateProgress progress, "Connecting to data source"
 
    ThisWorkbook.Unprotect "clausus"
    If Sheets("Input Sheet").AutoFilterMode Then Sheets("Koro").AutoFilter.ShowAllData
    
    frontend_regen.UnhideAllSheets

           
        Call extract_total.extract_input_sheet
        'modified by Noe, 31/03/2023

  progress = 60
    
    U_dialogbox.UpdateProgress progress, "Uploading data"
 
    Call UploadData_cf
 var.turn_very_hidden InputSheet, UserSelections
    ThisWorkbook.Protect "clausus"
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    progress = 100
    
    U_dialogbox.UpdateProgress progress, "Completed"
    Unload U_dialogbox
End Sub



