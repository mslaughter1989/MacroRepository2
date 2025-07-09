
Sub AutoProcessAndSaveFiles()
    Dim wb As Workbook, ws As Worksheet
    Dim sftpWb As Workbook, sftpWs As Worksheet
    Dim i As Long, j As Long
    Dim fileName As String, savePath As String, monthFolder As String
    Dim groupName As String, filePattern As String, fileDateFormat As String
    Dim matchFound As Boolean, fileDate As Date
    Dim zipColFound As Boolean, apexLogicApplied As Boolean
    Dim keywordList As Variant, keyword As Variant
    Dim zipFormattedList As String, apexAppliedList As String, untouchedList As String
    Dim regex As Object, matches As Object
    Dim dateStr As String, testStr As String
    Dim fullSavePath As String

    keywordList = Array("zip", "zipcode", "zip code", "postalcode", "postal code")
    zipFormattedList = ""
    apexAppliedList = ""
    untouchedList = ""

    ' Load SFTP mapping
    On Error Resume Next
    Set sftpWb = Workbooks.Open("C:\Users\MichaelSlaughter\AppData\Roaming\Microsoft\Excel\XLSTART\SFTPfiles.xlsx", ReadOnly:=True)
    On Error GoTo 0
    If sftpWb Is Nothing Then
        MsgBox "Could not open SFTPfiles.xlsx from XLSTART folder.", vbCritical
        Exit Sub
    End If
    Set sftpWs = sftpWb.Sheets("Sheet1")

    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True
    regex.IgnoreCase = True

    For Each wb In Application.Workbooks
        If wb.Name = sftpWb.Name Or wb.Name = "PERSONAL.XLSB" Then GoTo NextWorkbook
        If Not Right(wb.Name, 4) = ".csv" Then GoTo NextWorkbook
        If wb.Name = sftpWb.Name Then GoTo NextWorkbook
        If wb.Sheets.Count = 0 Then GoTo NextWorkbook
        Set ws = wb.Sheets(1)
        zipColFound = False: apexLogicApplied = False
        fileName = wb.Name

        ' === ZIP column formatting ===
        For j = 1 To ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
            If Not IsEmpty(ws.Cells(1, j)) Then
                testStr = LCase(Replace(Replace(Replace(ws.Cells(1, j).Text, "_", ""), "-", ""), " ", ""))
                For Each keyword In keywordList
                    If InStr(testStr, Replace(LCase(keyword), " ", "")) > 0 Then
                        ws.Columns(j).NumberFormat = "00000"
                        zipColFound = True
                        Exit For
                    End If
                Next keyword
            End If
        Next j

        ' === APEX special logic ===
        If InStr(UCase(fileName), "APEX") > 0 Then
            apexLogicApplied = True
            Dim dict As Object, rngP As Range, cell As Range
            Dim k As Long
            Set dict = CreateObject("Scripting.Dictionary")
            Set rngP = ws.Range("P2:P" & ws.Cells(ws.Rows.Count, "P").End(xlUp).Row)
            For Each cell In rngP
                If Not dict.Exists(cell.Value) Then dict.Add cell.Value, 1 Else dict(cell.Value) = dict(cell.Value) + 1
            Next cell
            For k = ws.Cells(ws.Rows.Count, "P").End(xlUp).Row To 2 Step -1
                If dict(ws.Cells(k, "P").Value) > 1 And ws.Cells(k, "N").Value <> "" Then ws.Rows(k).Delete
            Next k
            Set dict = CreateObject("Scripting.Dictionary")
            For Each cell In ws.Range("P2:P" & ws.Cells(ws.Rows.Count, "P").End(xlUp).Row)
                If Not dict.Exists(cell.Value) Then
                    dict.Add cell.Value, cell.Row
                Else
                    If ws.Cells(dict(cell.Value), "M").Value < ws.Cells(cell.Row, "M").Value Then
                        ws.Rows(dict(cell.Value)).Delete
                        dict(cell.Value) = cell.Row
                    Else
                        ws.Rows(cell.Row).Delete
                    End If
                End If
            Next cell
        End If

        ' === Match and Save Logic ===
        matchFound = False
        For i = 2 To sftpWs.Cells(sftpWs.Rows.Count, 1).End(xlUp).Row
            groupName = sftpWs.Cells(i, 1).Value
            filePattern = sftpWs.Cells(i, 2).Value
            savePath = sftpWs.Cells(i, 3).Value
            If InStr(fileName, Split(filePattern, "_")(0)) > 0 Then
                matchFound = True
                fileDateFormat = Split(Split(filePattern, "_")(UBound(Split(filePattern, "_"))), ".")(0)

                regex.Pattern = "\d+"
                Set matches = regex.Execute(fileName)
                For Each Match In matches
                    dateStr = Match.Value
                    If Len(dateStr) = 6 And fileDateFormat = "mmddyy" Then
                        fileDate = DateSerial(2000 + CInt(Right(dateStr, 2)), Left(dateStr, 2), Mid(dateStr, 3, 2))
                        Exit For
                    ElseIf Len(dateStr) = 8 And fileDateFormat = "mmddyyyy" Then
                        fileDate = DateSerial(CInt(Right(dateStr, 4)), Left(dateStr, 2), Mid(dateStr, 3, 2))
                        Exit For
                    ElseIf Len(dateStr) = 8 And fileDateFormat = "yyyymmdd" Then
                        fileDate = DateSerial(Left(dateStr, 4), Mid(dateStr, 5, 2), Right(dateStr, 2))
                        Exit For
                    End If
                Next

                monthFolder = Format(fileDate, "MM") & Format(fileDate, "MMM") & Right(Year(fileDate), 2)
                fullSavePath = savePath & "\" & monthFolder
                CreateFullPath fullSavePath

                If Dir(fullSavePath & "\" & fileName) <> "" Then
                    MsgBox "Duplicate file: " & fileName & " already exists in " & fullSavePath
                Else
                    wb.SaveAs fileName:=fullSavePath & "\" & fileName, FileFormat:=xlCSV
                End If
                Exit For
            End If
        Next i

        If zipColFound Then zipFormattedList = zipFormattedList & vbCrLf & "- " & fileName
        If apexLogicApplied Then apexAppliedList = apexAppliedList & vbCrLf & "- " & fileName
        If Not zipColFound And Not apexLogicApplied Then untouchedList = untouchedList & vbCrLf & "- " & fileName

NextWorkbook:
    Next wb

    sftpWb.Close False

    Dim msg As String
    msg = "Macro Completed" & vbCrLf & String(40, "-")
    If zipFormattedList <> "" Then msg = msg & vbCrLf & vbCrLf & "ZIP formatted:" & zipFormattedList
    If apexAppliedList <> "" Then msg = msg & vbCrLf & vbCrLf & "APEX processed:" & apexAppliedList
    If untouchedList <> "" Then msg = msg & vbCrLf & vbCrLf & "No changes:" & untouchedList
    MsgBox msg, vbInformation, "Macro Results"
End Sub
