Sub AutoProcessAndSaveFiles()
    Dim wb As Workbook, ws As Worksheet
    Dim rngP As Range, cell As Range
    Dim dict As Object
    Dim i As Long, j As Long
    Dim lastRow As Long, lastCol As Long
    Dim headerCell As Range, cleanHeader As String
    Dim zipColFound As Boolean, apexLogicApplied As Boolean
    Dim keywordList As Variant, keyword As Variant
    Dim zipFormattedList As String, apexAppliedList As String, untouchedList As String
    Dim fileName As String
    Dim sftpData As Variant, item As Variant
    Dim groupName As String, filePattern As String, savePath As String
    Dim matchFound As Boolean, fileDate As Date, dtString As String
    Dim monthFolder As String
    Dim regex As Object, matches As Object
    Dim tempName As String, tempPath As String, finalPath As String

    keywordList = Array("zip", "zipcode", "zip code", "postalcode", "postal code")
    zipFormattedList = ""
    apexAppliedList = ""
    untouchedList = ""

    MsgBox "Starting macro. Collecting embedded SFTP data..."
    sftpData = LoadSFTPData()
    MsgBox "Data loaded. Beginning workbook processing..."

    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True
    regex.IgnoreCase = True

    For Each wb In Application.Workbooks
        If wb.Sheets.Count = 0 Then GoTo NextWorkbook
        Set ws = wb.Sheets(1)
        zipColFound = False: apexLogicApplied = False
        fileName = wb.Name

        ' ZIP formatting
        lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        For j = 1 To lastCol
            cleanHeader = LCase(Replace(Replace(Replace(Trim(ws.Cells(1, j).Value), "_", ""), "-", ""), " ", ""))
            For Each keyword In keywordList
                If InStr(cleanHeader, Replace(LCase(keyword), " ", "")) > 0 Then
                    ws.Columns(j).NumberFormat = "00000"
                    zipColFound = True
                    Exit For
                End If
            Next keyword
        Next j

        ' APEX logic
        If InStr(UCase(fileName), "APEX") > 0 Then
            apexLogicApplied = True
            Set rngP = ws.Range("P2:P" & ws.Cells(ws.Rows.Count, "P").End(xlUp).Row)
            Set dict = CreateObject("Scripting.Dictionary")
            For Each cell In rngP
                If Not dict.Exists(cell.Value) Then dict.Add cell.Value, 1 Else dict(cell.Value) = dict(cell.Value) + 1
            Next cell
            For i = ws.Cells(ws.Rows.Count, "P").End(xlUp).Row To 2 Step -1
                If dict(ws.Cells(i, "P").Value) > 1 And ws.Cells(i, "N").Value <> "" Then ws.Rows(i).Delete
            Next i
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

        ' File matching
        matchFound = False
        For Each item In sftpData
            groupName = item(0): filePattern = item(1): savePath = item(2)
            If InStr(fileName, Split(filePattern, "_")(0)) > 0 Then
                matchFound = True
                If InStr(filePattern, "yyyymmdd") > 0 Then
                    regex.Pattern = "\d{8}"
                    Set matches = regex.Execute(fileName)
                    If matches.Count > 0 Then fileDate = DateSerial(Left(matches(0), 4), Mid(matches(0), 5, 2), Right(matches(0), 2))
                ElseIf InStr(filePattern, "mmddyyyy") > 0 Then
                    regex.Pattern = "\d{8}"
                    Set matches = regex.Execute(fileName)
                    If matches.Count > 0 Then fileDate = DateSerial(Right(matches(0), 4), Left(matches(0), 2), Mid(matches(0), 3, 2))
                ElseIf InStr(filePattern, "mmddyy") > 0 Then
                    regex.Pattern = "\d{6}"
                    Set matches = regex.Execute(fileName)
                    If matches.Count > 0 Then
                        dtString = matches(0)
                        fileDate = DateSerial("20" & Right(dtString, 2), Left(dtString, 2), Mid(dtString, 3, 2))
                    End If
                End If

                monthFolder = Format(fileDate, "MM") & Format(fileDate, "MMM") & Right(Year(fileDate), 2)
                Call CreateFullPath(savePath & "\" & monthFolder)
                If Dir(savePath & "\" & monthFolder & "\" & fileName) <> "" Then
                    MsgBox "Duplicate file: " & fileName & " already exists in " & monthFolder
                Else
                    wb.SaveAs fileName:=savePath & "\" & monthFolder & "\" & fileName, FileFormat:=xlCSV
                End If
                Exit For
            End If
        Next item

        If zipColFound Then zipFormattedList = zipFormattedList & vbCrLf & "- " & fileName
        If apexLogicApplied Then apexAppliedList = apexAppliedList & vbCrLf & "- " & fileName
        If Not zipColFound And Not apexLogicApplied Then untouchedList = untouchedList & vbCrLf & "- " & fileName

NextWorkbook:
    Next wb

    Dim msg As String
    msg = "Macro Complete!" & vbCrLf & String(40, "-")
    If zipFormattedList <> "" Then msg = msg & vbCrLf & vbCrLf & "ZIP columns formatted:" & zipFormattedList
    If apexAppliedList <> "" Then msg = msg & vbCrLf & vbCrLf & "APEX processed:" & apexAppliedList
    If untouchedList <> "" Then msg = msg & vbCrLf & vbCrLf & "No changes:" & untouchedList
    MsgBox msg, vbInformation, "Macro Results"
End Sub
