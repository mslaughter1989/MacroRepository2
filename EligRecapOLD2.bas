Attribute VB_Name = "EligRecapOLD2"
Sub EligRecapOLD2()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim eligibleWorkbooks As Collection
    Dim appliedWBs As Collection
    Dim skippedWBs As Collection
    Dim wbName As String
    Dim regEx As Object
    Dim wbItem As Variant
    Dim report As String

    Set eligibleWorkbooks = New Collection
    Set appliedWBs = New Collection
    Set skippedWBs = New Collection
    Set regEx = CreateObject("VBScript.RegExp")

    ' Regex to match: EligibilityRecapYYYY_MM_DD
    With regEx
        .Global = False
        .IgnoreCase = True
        .pattern = "^EligibilityRecap\d{4}_\d{2}_\d{2}"
    End With

    ' Identify eligible workbooks
    For Each wb In Application.Workbooks
        wbName = Left(wb.Name, InStrRev(wb.Name, ".") - 1) ' remove extension
        If regEx.Test(wbName) Then
            eligibleWorkbooks.Add wb
        Else
            skippedWBs.Add wb.Name
        End If
    Next wb

    ' Apply macro to each eligible workbook
    For Each wbItem In eligibleWorkbooks
        Set ws = wbItem.ActiveSheet
        Call Run_EligRecap_Filter(ws)
        appliedWBs.Add wbItem.Name
    Next wbItem

    ' Build report
    report = "APPLIED WORKBOOKS:" & vbCrLf
    For Each wbItem In appliedWBs
        report = report & " - " & wbItem & vbCrLf
    Next wbItem

    report = report & vbCrLf & "SKIPPED WORKBOOKS:" & vbCrLf
    For Each wbItem In skippedWBs
        report = report & " - " & wbItem & vbCrLf
    Next wbItem

    MsgBox report, vbInformation, "EligRecap Macro Report"
End Sub

' === FILTERING MACRO ===
Sub Run_EligRecap_Filter(ws As Worksheet)
    Dim lastRow As Long
    Dim i As Long
    Dim cellValC As String
    Dim cellValM As String
    Dim keepRow As Boolean

    ' Reset filters and unhide rows
    ws.AutoFilterMode = False
    ws.Rows.Hidden = False

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

    ' Step 1: Filter by Column C status
    For i = 2 To lastRow
        cellValC = ws.Cells(i, 3).Value
        If cellValC <> "Completed with Errors" And cellValC <> "Failed to Process File" Then
            ws.Rows(i).Hidden = True
        End If
    Next i

    ' Step 2: Further filter based on Column M errors
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

    ' Reapply AutoFilter and hide irrelevant columns
    ws.Rows("1:1").AutoFilter
    ws.Columns("C:C").EntireColumn.Hidden = True
    ws.Columns("E:E").EntireColumn.Hidden = True
    ws.Columns("I:L").EntireColumn.Hidden = True
    ws.Columns("N:O").EntireColumn.Hidden = True
End Sub


