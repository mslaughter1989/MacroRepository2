Attribute VB_Name = "DOBCleaner"
Sub DOBCleaner()
    Dim ws As Worksheet
    Dim cell As Range
    Dim headerCell As Range
    Dim dobColumn As Long
    Dim regEx As Object
    Dim cleanHeader As String
    Dim i As Long
    Dim headerRow As Range
    Dim pattern As String
    
    Set ws = ActiveSheet
    Set headerRow = ws.Rows(1)
    dobColumn = 0

    ' Create regex to detect mm/dd/yyyy format
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.pattern = "^(0[1-9]|1[0-2])/(0[1-9]|[12]\d|3[01])/\d{4}$"
    regEx.IgnoreCase = False
    regEx.Global = False

    ' Loop through header row to find DOB column
    For Each headerCell In headerRow.Cells
        cleanHeader = UCase(headerCell.Value)
        cleanHeader = Replace(cleanHeader, ".", "")
        cleanHeader = Replace(cleanHeader, "/", "")
        cleanHeader = Replace(cleanHeader, "-", "")
        cleanHeader = Replace(cleanHeader, "_", "")
        cleanHeader = Replace(cleanHeader, " ", "")
        
        If InStr(cleanHeader, "DATEOFBIRTH") > 0 Or InStr(cleanHeader, "DOB") > 0 Then
            dobColumn = headerCell.Column
            Exit For
        End If
    Next headerCell

    If dobColumn = 0 Then
        MsgBox "No column matching 'Date of Birth' was found.", vbExclamation
        Exit Sub
    End If

    ' Loop through each cell in the DOB column (starting at row 2)
    For i = 2 To ws.Cells(ws.Rows.count, dobColumn).End(xlUp).Row
        With ws.Cells(i, dobColumn)
            If Not regEx.Test(.Text) Then
                .Value = "" ' Remove non-standard formatted entries
            End If
        End With
    Next i

    MsgBox "Date of Birth column cleaned successfully.", vbInformation
End Sub

