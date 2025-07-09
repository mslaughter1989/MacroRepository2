Sub CreateFullPath(fullPath As String)
    Dim fso As Object, pathParts() As String
    Dim curPath As String
    Dim i As Integer
    Set fso = CreateObject("Scripting.FileSystemObject")

    pathParts = Split(fullPath, "\")
    curPath = pathParts(0)

    For i = 1 To UBound(pathParts)
        curPath = curPath & "\" & pathParts(i)
        If Not fso.FolderExists(curPath) Then fso.CreateFolder curPath
    Next i
End Sub
