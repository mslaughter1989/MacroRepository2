Attribute VB_Name = "FarWestDataClean"
Sub FarWestDataClean()
    Dim ws As Worksheet
    Dim i As Long, lastRow As Long

    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row

    Dim problemIDs As Variant
    problemIDs = Array("F678092834", "F658325147", "F607635887", _
                       "F627103182", "F603789034", "F604770423", _
                       "F612531357", "F760330841")

    For i = 2 To lastRow
        If UBound(Filter(problemIDs, ws.Cells(i, "C").Value)) >= 0 Then
            ' Shift City-State-Zip data rightward correctly
            If ws.Cells(i, "N").Value Like "* CA*" Or ws.Cells(i, "N").Value Like "* NV*" Or ws.Cells(i, "N").Value Like "* CO*" Then
                ws.Range(ws.Cells(i, "N"), ws.Cells(i, "U")).Insert Shift:=xlToRight
            End If

            ' Fix Zip code data alignment
            If Not IsNumeric(ws.Cells(i, "P").Value) Then
                ws.Range(ws.Cells(i, "P"), ws.Cells(i, "U")).Insert Shift:=xlToRight
            End If

            ' Align email addresses properly
            If ws.Cells(i, "Q").Value Like "*@*" Then
                ws.Cells(i, "P").Value = ws.Cells(i, "Q").Value
                ws.Cells(i, "Q").ClearContents
            End If

            ' Align phone numbers properly
            If ws.Cells(i, "R").Value Like "*-*-*" Then
                ws.Cells(i, "Q").Value = ws.Cells(i, "R").Value
                ws.Cells(i, "R").ClearContents
            End If

            ' Correct location code
            If ws.Cells(i, "U").Value <> "" Then
                ws.Cells(i, "R").Value = ws.Cells(i, "U").Value
                ws.Cells(i, "U").ClearContents
            End If

            ' Remove trailing columns labeled "Unnamed"
            Dim col As Integer
            For col = ws.Cells(1, Columns.count).End(xlToLeft).Column To 1 Step -1
                If ws.Cells(1, col).Value Like "Unnamed*" Then
                    ws.Columns(col).Delete
                End If
            Next col
        End If
    Next i

    MsgBox "Persistent Data Misalignment Corrected!", vbInformation
End Sub



