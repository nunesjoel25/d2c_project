Attribute VB_Name = "frontend_regen"
Option Explicit

Sub copyRange()
 Dim frontendRegen As New frontendRegeneration
    
    Set frontendRegen.copyRange_input = Sheets("InputCellDefination").Range("InputCellDefination")
    Set frontendRegen.copyRange_koro = Sheets("KoroSheetDefination").Range("KoroSheetDefination")
    Set frontendRegen.pasteRange_input = Sheets("Koro").Range("F7")
    Set frontendRegen.pasteRange_koro = Sheets("Input Sheet").Range("I8")
    Set frontendRegen.copyRange_non_key = Sheets("Non-Key").Range("NonKeySheetCellDefination")
    Set frontendRegen.pasteRange_non_key = Sheets("Non-Key").Range("F45")
        
    
    frontendRegen.combinations_koro = Sheets("Combinations").Range("B2").value
    frontendRegen.combinations_input = Sheets("material_combinations").Range("B2").value

    
  
    Application.Calculation = xlCalculationManual
    
    If frontendRegen.UserSelectionLevel = "SeAG" And frontendRegen.UserSelectionIndicator = "Key" Then
        frontendRegen.CopyPaste frontendRegen.copyRange_input, frontendRegen.pasteRange_input, frontendRegen.combinations_koro
        
        ElseIf frontendRegen.UserSelectionLevel = "SeAG" And frontendRegen.UserSelectionIndicator = "Non-Key" Then
         
        frontendRegen.CopyPaste frontendRegen.copyRange_non_key, frontendRegen.pasteRange_non_key, frontendRegen.combinations_koro - 1
    
        ElseIf frontendRegen.UserSelectionLevel = "Sales Organisation" Then
            
        frontendRegen.CopyPaste frontendRegen.copyRange_koro, frontendRegen.pasteRange_koro, frontendRegen.combinations_input
End If
    

End Sub

Sub MaterialListcopy()
    Dim frontendRegen As New frontendRegeneration
    Dim sourceSheet_koro As Worksheet
    Dim targetSheet_koro As Worksheet
    Dim sourceTable_koro As ListObject
    Dim sourceSheet_koro_non_key As Worksheet
    Dim targetSheet_koro_non_key As Worksheet
    Dim sourceTable_koro_non_key As ListObject
    Dim targetCell_koro As Range
    Dim targetCell_koro_non_key As Range
    Dim sourceSheet_input As Worksheet
    Dim targetSheet_input As Worksheet
    Dim sourceTable_input As ListObject
    Dim targetCell_input As Range
    Dim UserSelectionLevel As String
    Dim UserSelection_indicator As String
    'Set the source and target sheets
    
    Set sourceSheet_koro = Worksheets("List")
    Set targetSheet_koro = Worksheets("Koro")
    Set sourceSheet_koro_non_key = Worksheets("nonkey_list")
    Set targetSheet_koro_non_key = Worksheets("Non-Key")
    Set sourceSheet_input = Worksheets("material_list")
    Set targetSheet_input = Worksheets("Input Sheet")
    'Set the source table using the ListObjects property
    
    Set sourceTable_koro = sourceSheet_koro.ListObjects("List")
    Set sourceTable_input = sourceSheet_input.ListObjects("material_list")
    Set sourceTable_koro_non_key = sourceSheet_koro_non_key.ListObjects("nonkey_list")
    UserSelectionLevel = Worksheets("User Selections").Range("C6").value
  
    Set targetCell_koro = targetSheet_koro.Range("C7")
    Set targetCell_input = targetSheet_input.Range("C8")
    Set targetCell_koro_non_key = targetSheet_koro_non_key.Range("C7")
    
    
    If frontendRegen.UserSelectionLevel = "Sales Organisation" Then
    sourceTable_input.DataBodyRange.Copy Destination:=targetCell_input

    
    ElseIf frontendRegen.UserSelectionLevel = "SeAG" Then
        
            If frontendRegen.UserSelectionIndicator = "Key" Then
            
    sourceTable_koro.DataBodyRange.Copy Destination:=targetCell_koro
    
    Else
    
    sourceTable_koro_non_key.DataBodyRange.Copy Destination:=targetCell_koro_non_key
    End If
End If
    
End Sub


Sub copy_tables()
Dim sheet As New frontendRegeneration
Set sheet.sourceSheet_traffic = Worksheets("Traffic Actuals_count")
Set sheet.sourceSheet_orders = Worksheets("Orders(SAP Hybris_Material)_cou")
Set sheet.targetSheet_traffic = Worksheets("PDP ACT")
Set sheet.targetSheet_orders = Worksheets("MATERIAL ACT")

Set sheet.sourceTable_traffic = sheet.sourceSheet_traffic.ListObjects("Traffic_Actuals_country_selector")
Set sheet.sourceTable_quantity = sheet.sourceSheet_orders.ListObjects("Orders_SAP_Hybris_Material__country_selector")
    
    Set sheet.targetCell_traffic = sheet.targetSheet_traffic.Range("A1")
    sheet.sourceTable_traffic.DataBodyRange.Copy Destination:=sheet.targetCell_traffic
        
    Set sheet.targetCell_orders = sheet.targetSheet_orders.Range("A1")
    sheet.sourceTable_quantity.DataBodyRange.Copy Destination:=sheet.targetCell_orders
    
End Sub
Sub UnhideAllSheets()
    Dim ws As Worksheet
    
    For Each ws In ActiveWorkbook.Worksheets
        ws.Visible = xlSheetVisible
    Next ws
    
UploadConfig.Visible = xlSheetVeryHidden
End Sub
Sub delete_overrides_key()
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
 Dim ws As Worksheet
 Dim lastRow As Long

    Set ws = ThisWorkbook.Worksheets("Koro")

    lastRow = ws.Cells(ws.Rows.count, "J").End(xlUp).Row
    Dim myRange As Range
    Set myRange = ws.Range("K6:AD" & lastRow)
    Dim i As Integer
    For i = myRange.Rows.count To 1 Step -1
        If Range("J" & i + 5).value = "Sales Quantity Override" Then
            myRange.Rows(i).clearContents
        End If
    Next i
Application.ScreenUpdating = False
Application.Calculation = xlCalculationAutomatic
End Sub
Sub delete_overrides_key_current_month()
 Dim ws As Worksheet
 Dim searchRange As Range
 Dim lastRow As Long
 Dim current_month As Long
 Dim columnIndex As Integer
 Dim columnLetter As String
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
    Set ws = ThisWorkbook.Worksheets("Koro")
    current_month = Sheets("Settings").Range("L3")
    Set searchRange = ws.Rows(6)
    
    columnIndex = Application.Match(current_month, searchRange, 0)
    columnLetter = Split(Cells(1, columnIndex).Address, "$")(1)
  
    
    lastRow = ws.Cells(ws.Rows.count, "J").End(xlUp).Row
    Dim myRange As Range
    Set myRange = ws.Range(columnLetter & "6:" & columnLetter & lastRow)
    Dim i As Integer
    For i = myRange.Rows.count To 1 Step -1
        If Range("J" & i + 5).value = "Sales Quantity Override" Then
            myRange.Rows(i).clearContents
        End If
    Next i
Application.ScreenUpdating = False
Application.Calculation = xlCalculationAutomatic
End Sub
Sub GroupColumns_input_sheet()
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

Sub formatSheet(ws As Worksheet)
    Dim var As New frontendRegeneration
    Dim groupRowsInput As New frontendRegeneration
    Dim lastRow As Long
    Dim rng_koro As Range
    Dim rng_input As Range
    Dim rng_non_key As Range
      
    lastRow = ws.Cells(Rows.count, "C").End(xlUp).Row
    
    Set rng_koro = KoroSheet.Range("F7:F" & lastRow)
    Set rng_input = InputSheet.Range("I1:I" & lastRow)
    Set rng_non_key = NonKey.Range("F1:F" & lastRow)
    
    Select Case var.UserSelectionLevel
        
        Case Is = "Sales Organisation"
            InputSheet.Activate
            InputSheet.Range("M8").Select
            var.FreezePanes InputSheet.Range("M8")
            var.formatrows InputSheet, 6
            var.groupRows rng_input, InputSheet
            var.u_hide_columns_rows InputSheet, True
        
        Case Is = "SeAG"
            Select Case var.UserSelectionIndicator
                
                Case Is = "Key"
                    KoroSheet.Activate
                    KoroSheet.Range("K7").Select
                    var.FreezePanes KoroSheet.Range("K7")
                    var.formatrows KoroSheet, 9
                    var.groupRows rng_koro, KoroSheet
                    var.u_hide_columns_rows KoroSheet, True
            
                
                Case Is = "Non-Key"
                    NonKey.Activate
                    NonKey.Range("K7").Select
                    var.FreezePanes NonKey.Range("K7")
                    var.formatrows NonKey, 6
                    var.groupRows rng_non_key, NonKey
                    var.u_hide_columns_rows NonKey, True
            End Select
    End Select
End Sub
Sub clearAll()
Dim clear As New frontendRegeneration

    Select Case clear.UserSelectionLevel
    Case Is = "Sales Organisation"
        clear.clearAll Range("KoroSheetGrid"), InputSheet, 1
        clear.u_hide_columns_rows InputSheet, False
    Case Is = "SeAG"
        Select Case clear.UserSelectionIndicator
            Case Is = "Key"
                clear.clearAll Range("input_grid_key"), KoroSheet, 1
                clear.u_hide_columns_rows KoroSheet, False
                pdpAct.Cells.clear
                materialAct.Cells.clear
            Case Is = "Non-Key"
                clear.clearAll Range("NonKeyCellDefination"), NonKey, 0
                clear.u_hide_columns_rows NonKey, False
                pdpAct.Cells.clear
                materialAct.Cells.clear
        End Select
    End Select
End Sub

Sub RefreshQuery()
Dim query_refresh As New frontendRegeneration
Dim progress As Integer
    Select Case query_refresh.UserSelectionLevel
    Case Is = "Sales Organisation"
        query_refresh.QueryRefresh 9
    Case Is = "SeAG"
        Select Case query_refresh.UserSelectionIndicator
            Case Is = "Key"
            progress = 25
            U_dialogbox.UpdateProgress progress, "Refreshing sub data queries.."
        query_refresh.QueryRefresh 1
            progress = 45
            U_dialogbox.UpdateProgress progress, "Refreshing Order and Traffic data queries.."
        query_refresh.QueryRefresh 3
            Case Is = "Non-Key"
            progress = 25
            U_dialogbox.UpdateProgress progress, "Refreshing sub data queries.."
        query_refresh.QueryRefresh 1
            progress = 30
            U_dialogbox.UpdateProgress progress, "Refreshing Orders and Traffic data queries.."
        query_refresh.QueryRefresh 3
            progress = 45
            U_dialogbox.UpdateProgress progress, "Running MA Calcs..."
        query_refresh.QueryRefresh 11
        End Select
    End Select
End Sub
        
Sub btns_onAction(control As IRibbonControl)
    Select Case control.ID
       
        Case "Button1"
        Call ip_extract.refreshfromServer
        Case "Button2"
            ip_extract.UploadToServer
        Case "Button3"
            ip_extract.ip_extract_call
        Case "Button4"
            refresh_frontend.refresh_total
        Case "Button5"
            refresh_frontend.on_demand_refresh
        Case "Button6"
            refresh_frontend.clear_filters
        Case "Button8"
        refresh_frontend.training_link
        Case "Button10"
        retrive_data.retrive_data
        Case "Button11"
            frontend_regen.delete_overrides_key_current_month
        Case "Button12"
            frontend_regen.delete_overrides_key
     
     End Select
End Sub
        
Sub try()
Dim var As New frontendRegeneration
 var.turn_very_hidden KoroSheet, TotalSheet, UserSelections
End Sub

