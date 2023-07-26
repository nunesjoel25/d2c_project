Attribute VB_Name = "export_modules"
Sub ExportModules()
    Dim VBProj As VBIDE.VBProject
    Dim VBComp As VBIDE.VBComponent
    Dim FolderPath As String
    
    ' Set the folder path where you want to save the modules
    FolderPath = "C:\Users\7000029397\Downloads\D2C project\"
    
    Set VBProj = ActiveWorkbook.VBProject
    For Each VBComp In VBProj.VBComponents
        If VBComp.Type = vbext_ct_StdModule Or VBComp.Type = vbext_ct_ClassModule Then
            VBComp.Export FolderPath & VBComp.Name & IIf(VBComp.Type = vbext_ct_StdModule, ".bas", ".cls")
        End If
    Next VBComp
End Sub

