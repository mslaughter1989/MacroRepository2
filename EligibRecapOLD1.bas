Attribute VB_Name = "EligibRecapOLD1"
Sub EligibRecapOLD1()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim cellValC As String
    Dim cellValM As String
    Dim keepRow As Boolean
    Set ws = ActiveSheet

    ' Reset filters and unhide rows
    ws.AutoFilterMode = False
    ws.Rows.Hidden = False

    ' Find last used row based on column A and M
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    If ws.Cells(ws.Rows.count, 13).End(xlUp).Row > lastRow Then
        lastRow = ws.Cells(ws.Rows.count, 13).End(xlUp).Row
    End If

    ' Sort Column A alphabetically
    With ws.Sort
        .SortFields.Clear
        .SortFields.Add2 key:=ws.Range("A2:A" & lastRow), _
            SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SetRange ws.Range("A1:O" & lastRow)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    ' Step 1: Hide all rows that do NOT match required Column C values
    For i = 2 To lastRow
        cellValC = ws.Cells(i, 3).Value
        If cellValC <> "Completed with Errors" And cellValC <> "Failed to Process File" Then
            ws.Rows(i).Hidden = True
        End If
    Next i

    ' Step 2: Further hide rows (already visible) that do NOT contain valid Column M values
    For i = 2 To lastRow
        If ws.Rows(i).Hidden = False Then
            cellValM = ws.Cells(i, 13).Value
            keepRow = False

            If InStr(1, cellValM, "Duplicate CMID for unique CMID FileProcess", vbTextCompare) > 0 Then keepRow = True
            If InStr(1, cellValM, "Invalid Product Offering", vbTextCompare) > 0 Then keepRow = True
            If InStr(1, cellValM, "Invalid Group ID", vbTextCompare) > 0 Then keepRow = True
            If Trim(cellValM) = "" Then keepRow = True

            If Not keepRow Then
                ws.Rows(i).Hidden = True
            End If
        End If
    Next i

    ' Reapply autofilter on header row
    ws.Rows("1:1").AutoFilter

    ' Hide irrelevant columns
    ws.Columns("C:C").EntireColumn.Hidden = True
    ws.Columns("E:E").EntireColumn.Hidden = True
    ws.Columns("I:L").EntireColumn.Hidden = True
    ws.Columns("N:O").EntireColumn.Hidden = True
End Sub

