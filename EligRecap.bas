Attribute VB_Name = "EligRecap"
Sub EligRecap()
    Dim wb As Workbook, ws As Worksheet
    Dim eligibleWorkbooks As Collection, appliedWBs As Collection, skippedWBs As Collection
    Dim wbName As String, regEx As Object, wbItem As Variant
    Dim report As String, masterWB As Workbook, masterWS As Worksheet
    Dim lastRowDest As Long, isFirstCopy As Boolean
    Dim savePath As String, timeStamp As String
    Dim checkMark As String

    Set eligibleWorkbooks = New Collection
    Set appliedWBs = New Collection
    Set skippedWBs = New Collection
    Set regEx = CreateObject("VBScript.RegExp")
    Set masterWB = Workbooks.Add
    Set masterWS = masterWB.Sheets(1)
    masterWS.Name = "Combined EligRecap"
    isFirstCopy = True

    checkMark = ChrW(&H2713) ' Unicode checkmark ?

    ' Timestamp and path
    timeStamp = Format(Now, "yyyymmdd_HHmm")
    savePath = "C:\Users\MichaelSlaughter\Downloads\EligibilityRecap_CombinedResults_" & timeStamp & ".xlsx"

    ' Match pattern: EligibilityRecapYYYY_MM_DD
    With regEx
        .Global = False
        .IgnoreCase = True
        .pattern = "^EligibilityRecap\d{4}_\d{2}_\d{2}"
    End With

    ' Identify matching workbooks
    For Each wb In Application.Workbooks
        If InStrRev(wb.Name, ".") > 0 Then
            wbName = Left(wb.Name, InStrRev(wb.Name, ".") - 1)
        Else
            wbName = wb.Name
        End If

        If regEx.Test(wbName) Then
            eligibleWorkbooks.Add wb
        Else
            skippedWBs.Add wb.Name
        End If
    Next wb

    ' Process each eligible workbook
    For Each wbItem In eligibleWorkbooks
        Set ws = wbItem.ActiveSheet
        Call Run_EligRecap_Filter(ws)
        appliedWBs.Add wbItem.Name

        ' Copy visible filtered rows
        With ws.AutoFilter.Range
            If isFirstCopy Then
                .SpecialCells(xlCellTypeVisible).Copy ' Include headers
                isFirstCopy = False
            Else
                .Offset(1, 0).Resize(.Rows.count - 1).SpecialCells(xlCellTypeVisible).Copy ' Skip header
            End If
        End With

        ' Paste into master sheet
        With masterWS
            If Application.WorksheetFunction.CountA(.Rows(1)) = 0 Then
                lastRowDest = 1
            Else
                lastRowDest = .Cells(.Rows.count, 1).End(xlUp).Row + 1
            End If
            .Cells(lastRowDest, 1).PasteSpecial Paste:=xlPasteValues
        End With
    Next wbItem

    Application.CutCopyMode = False

    ' Sort column A alphabetically (A-Z)
    With masterWS
        .Range("A1").CurrentRegion.Sort Key1:=.Range("A1"), Order1:=xlAscending, Header:=xlYes
    End With

    ' Save but keep open
    masterWB.SaveAs Filename:=savePath, FileFormat:=xlOpenXMLWorkbook

    ' Build summary report
    report = "APPLIED WORKBOOKS:" & vbCrLf
    For Each wbItem In appliedWBs
        report = report & " - " & wbItem & vbCrLf
    Next wbItem

    report = report & vbCrLf & "SKIPPED WORKBOOKS:" & vbCrLf
    For Each wbItem In skippedWBs
        report = report & " - " & wbItem & vbCrLf
    Next wbItem

    report = report & vbCrLf & checkMark & " Combined file saved to:" & vbCrLf & savePath & vbCrLf & vbCrLf & "It has been left open for your review."

    MsgBox report, vbInformation, "EligRecap Macro Report"
End Sub

' === FILTERING LOGIC ===
Sub Run_EligRecap_Filter(ws As Worksheet)
    Dim lastRow As Long, i As Long
    Dim cellValC As String, cellValM As String
    Dim keepRow As Boolean

    ws.AutoFilterMode = False
    ws.Rows.Hidden = False

    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    If ws.Cells(ws.Rows.count, 13).End(xlUp).Row > lastRow Then
        lastRow = ws.Cells(ws.Rows.count, 13).End(xlUp).Row
    End If

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

    ' Filter by Column C
    For i = 2 To lastRow
        cellValC = ws.Cells(i, 3).Value
        If cellValC <> "Completed with Errors" And cellValC <> "Failed to Process File" Then
            ws.Rows(i).Hidden = True
        End If
    Next i

    ' Then filter by Column M
    For i = 2 To lastRow
        If ws.Rows(i).Hidden = False Then
            cellValM = ws.Cells(i, 13).Value
            keepRow = False
            If InStr(1, cellValM, "Duplicate CMID for unique CMID FileProcess", vbTextCompare) > 0 Then keepRow = True
            If InStr(1, cellValM, "Invalid Product Offering", vbTextCompare) > 0 Then keepRow = True
            If InStr(1, cellValM, "Invalid Group ID", vbTextCompare) > 0 Then keepRow = True
            If Trim(cellValM) = "" Then keepRow = True
            If Not keepRow Then ws.Rows(i).Hidden = True
        End If
    Next i

    ' Final clean-up
    ws.Rows("1:1").AutoFilter
    ws.Columns("C:C").EntireColumn.Hidden = True
    ws.Columns("E:E").EntireColumn.Hidden = True
    ws.Columns("I:L").EntireColumn.Hidden = True
    ws.Columns("N:O").EntireColumn.Hidden = True
End Sub


