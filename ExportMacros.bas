Attribute VB_Name = "ExportMacros"
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Dim vbComp As Object
    Dim exportPath As String

    ' Set the export folder
    exportPath = "C:\Users\MichaelSlaughter\Excel-Macros\"

    ' Loop through each module in the VBA project
    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        ' Export only standard code modules (bas)
        Select Case vbComp.Type
            Case 1, 2, 3 ' Module, Class Module, UserForm
                On Error Resume Next
                vbComp.Export exportPath & vbComp.Name & ".bas"
                On Error GoTo 0
        End Select
    Next vbComp
End Sub

