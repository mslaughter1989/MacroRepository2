Attribute VB_Name = "FormatZip"
Sub FormatZip()
Attribute FormatZip.VB_ProcData.VB_Invoke_Func = "q\n14"
    Dim ws As Worksheet
    Dim headerCell As Range
    Dim zipCol As Range
    Dim headerRow As Range
    Dim lastCol As Long
    Dim keywordList As Variant
    Dim cell As Range
    Dim i As Long
    Dim cleanHeader As String

    ' Define common ZIP code header keywords
    keywordList = Array("zip", "zipcode", "zip code", "postalcode", "postal code")

    Set ws = ActiveSheet
    Set headerRow = ws.Rows(1)
    lastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column

    ' Loop through each column in the first row
    For i = 1 To lastCol
        Set headerCell = ws.Cells(1, i)
        cleanHeader = LCase(Replace(Replace(Replace(Trim(headerCell.Value), "_", ""), "-", ""), " ", "")) ' Remove underscores, dashes, spaces

        For Each keyword In keywordList
            If InStr(cleanHeader, Replace(LCase(keyword), " ", "")) > 0 Then
                Set zipCol = ws.Columns(i)
                Exit For
            End If
        Next keyword

        If Not zipCol Is Nothing Then Exit For
    Next i

    ' Format found column
    If Not zipCol Is Nothing Then
        zipCol.NumberFormat = "00000"
        MsgBox "ZIP code column found and formatted: " & zipCol.Address, vbInformation
    Else
        MsgBox "No ZIP code column found based on known patterns.", vbExclamation
    End If
End Sub
Private Sub Workbook_Open()
    Application.OnKey "^q", "FormatZipCodeColumn"
End Sub

