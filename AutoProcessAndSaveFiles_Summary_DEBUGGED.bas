
Attribute VB_Name = "AutoProcessAndSaveFiles_Summary"
Option Explicit

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
    Dim savedFiles As String, skippedFiles As String
    Dim fileName As String, fileDate As Date, dateStr As String
    Dim sftpData As Variant, item As Variant
    Dim groupName As String, filePattern As String, savePath As String, fileDateFormat As String
    Dim matchFound As Boolean, monthFolder As String, fullSavePath As String
    Dim regex As Object, matches As Object

    keywordList = Array("zip", "zipcode", "zip code", "postalcode", "postal code")
    zipFormattedList = "": apexAppliedList = "": untouchedList = ""
    savedFiles = "": skippedFiles = ""

    ' Load mapping from SFTPfiles.xlsx
    sftpData = LoadSFTPMappingFromSheet()

    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True: regex.IgnoreCase = True

    For Each wb In Application.Workbooks
        If wb.Path = "" Or UCase(wb.Name) = "PERSONAL.XLSB" Then GoTo NextWorkbook
        If wb.Sheets.Count = 0 Then GoTo NextWorkbook
        Set ws = wb.Sheets(1)

        fileName = wb.Name
        zipColFound = False: apexLogicApplied = False

        ' === ZIP formatting ===
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

        ' === APEX logic ===
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
        End If

        ' === File matching and saving ===
        matchFound = False
        For Each item In sftpData
            groupName = item(0): filePattern = item(1): savePath = item(2)
            If InStr(fileName, Split(filePattern, "_")(0)) > 0 Then
                matchFound = True
                fileDateFormat = ExtractDateFormat(filePattern)
                regex.Pattern = DatePattern(fileDateFormat)
                Set matches = regex.Execute(fileName)
                If matches.Count > 0 Then
                    dateStr = matches(0)
                    fileDate = ConvertToDate(dateStr, fileDateFormat)
                Else
                    GoTo SkipFile
                End If

                monthFolder = Format(fileDate, "MM") & Format(fileDate, "MMM") & Right(Year(fileDate), 2)
                fullSavePath = savePath & "" & monthFolder
                CreateFullPath fullSavePath

                If Dir(fullSavePath & "" & fileName) <> "" Then
                    MsgBox "Duplicate file exists. Skipping save: " & fileName
                    GoTo SkipFile
                End If

                wb.SaveAs fileName:=fullSavePath & "" & fileName, FileFormat:=xlCSV
                savedFiles = savedFiles & vbCrLf & fileName & " → " & fullSavePath
                GoTo FinalizeFile
            End If
        Next item

SkipFile:
        skippedFiles = skippedFiles & vbCrLf & fileName

FinalizeFile:
        If zipColFound Then zipFormattedList = zipFormattedList & vbCrLf & "- " & fileName
        If apexLogicApplied Then apexAppliedList = apexAppliedList & vbCrLf & "- " & fileName
        If Not zipColFound And Not apexLogicApplied Then untouchedList = untouchedList & vbCrLf & "- " & fileName

NextWorkbook:
    Next wb

    Dim msg As String
    msg = "Macro Completed" & vbCrLf & String(40, "-")
    If savedFiles <> "" Then msg = msg & vbCrLf & vbCrLf & "Saved files:" & savedFiles
    If skippedFiles <> "" Then msg = msg & vbCrLf & vbCrLf & "Skipped files:" & skippedFiles
    If zipFormattedList <> "" Then msg = msg & vbCrLf & vbCrLf & "ZIP formatted:" & zipFormattedList
    If apexAppliedList <> "" Then msg = msg & vbCrLf & vbCrLf & "APEX processed:" & apexAppliedList
    If untouchedList <> "" Then msg = msg & vbCrLf & vbCrLf & "No changes:" & untouchedList
    MsgBox msg, vbInformation, "Macro Summary"
End Sub
