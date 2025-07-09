Attribute VB_Name = "Module1"
Sub ExportAllVBAModules()
    Dim vbComp As Object
    Dim exportFolder As String

    ' Define export folder
    exportFolder = "C:\Macro_Repository\"
    If Dir(exportFolder, vbDirectory) = "" Then MkDir exportFolder

    ' Loop through all VBA components and export them
    For Each vbComp In Application.Workbooks("PERSONAL1.XLSB").VBProject.VBComponents
        Select Case vbComp.Type
            Case 1: vbComp.Export exportFolder & vbComp.Name & ".bas"  ' Standard module
            Case 2: vbComp.Export exportFolder & vbComp.Name & ".cls"  ' Class module
            Case 3: vbComp.Export exportFolder & vbComp.Name & ".frm"  ' UserForm
        End Select
    Next vbComp

    MsgBox "All modules exported to: " & exportFolder, vbInformation
End Sub

