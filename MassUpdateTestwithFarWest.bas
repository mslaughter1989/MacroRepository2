Attribute VB_Name = "MassUpdateTestwithFarWest"
Sub MassUpdateTestwithFarWest()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim rngP As Range, cell As Range
    Dim dict As Object
    Dim i As Long, j As Long, lastRow As Long, lastCol As Long
    Dim headerCell As Range
    Dim cleanHeader As String
    Dim zipColFound As Boolean, apexLogicApplied As Boolean, arWestLogicApplied As Boolean
    Dim keywordList As Variant, keyword As Variant

    Dim zipFormattedList As String
    Dim apexAppliedList As String
    Dim arWestAppliedList As String
    Dim untouchedList As String

    ' ZIP header keywords
    keywordList = Array("zip", "zipcode", "zip code", "postalcode", "postal code")
    zipFormattedList = ""
    apexAppliedList = ""
    arWestAppliedList = ""
    untouchedList = ""

    For Each wb In Application.Workbooks
        If wb.Sheets.count > 0 Then
            Set ws = wb.Sheets(1)
            zipColFound = False
            apexLogicApplied = False
            arWestLogicApplied = False

            ' === ZIP code formatting ===
            lastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column

            For j = 1 To lastCol
                Set headerCell = ws.Cells(1, j)
                cleanHeader = LCase(Replace(Replace(Replace(Trim(headerCell.Value), "_", ""), "-", ""), " ", ""))

                For Each keyword In keywordList
                    If InStr(cleanHeader, Replace(LCase(keyword), " ", "")) > 0 Then
                        ws.Columns(j).NumberFormat = "00000"
                        zipColFound = True
                        Exit For
                    End If
                Next keyword
            Next j

            ' === APEX logic ===
            If InStr(1, UCase(wb.Name), "APEX") > 0 Then
                apexLogicApplied = True
                Set rngP = ws.Range("P2:P" & ws.Cells(ws.Rows.count, "P").End(xlUp).Row)
                Set dict = CreateObject("Scripting.Dictionary")

                For Each cell In rngP
                    If Not dict.Exists(cell.Value) Then
                        dict.Add cell.Value, 1
                    Else
                        dict(cell.Value) = dict(cell.Value) + 1
                    End If
                Next cell

                For i = ws.Cells(ws.Rows.count, "P").End(xlUp).Row To 2 Step -1
                    If dict.Exists(ws.Cells(i, "P").Value) And dict(ws.Cells(i, "P").Value) > 1 _
                        And ws.Cells(i, "N").Value <> "" Then
                        ws.Rows(i).Delete
                    End If
                Next i

                Set dict = CreateObject("Scripting.Dictionary")
                lastRow = ws.Cells(ws.Rows.count, "P").End(xlUp).Row

                For Each cell In ws.Range("P2:P" & lastRow)
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

            ' === arWestRestaurant_807303 logic ===
            If wb.Name Like "arWestRestaurant_807303_########*" Then
                arWestLogicApplied = True
                lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
                Dim problemIDs As Variant
                problemIDs = Array("F678092834", "F658325147", "F607635887", _
                                   "F627103182", "F603789034", "F604770423", _
                                   "F612531357", "F760330841")

                For i = 2 To lastRow
                    If UBound(Filter(problemIDs, ws.Cells(i, "C").Value)) >= 0 Then
                        ' Shift City-State-Zip
                        If ws.Cells(i, "N").Value Like "* CA*" Or ws.Cells(i, "N").Value Like "* NV*" Or ws.Cells(i, "N").Value Like "* CO*" Then
                            ws.Range(ws.Cells(i, "N"), ws.Cells(i, "U")).Insert Shift:=xlToRight
                        End If
                        ' Fix Zip code
                        If Not IsNumeric(ws.Cells(i, "P").Value) Then
                            ws.Range(ws.Cells(i, "P"), ws.Cells(i, "U")).Insert Shift:=xlToRight
                        End If
                        ' Email alignment
                        If ws.Cells(i, "Q").Value Like "*@*" Then
                            ws.Cells(i, "P").Value = ws.Cells(i, "Q").Value
                            ws.Cells(i, "Q").ClearContents
                        End If
                        ' Phone alignment
                        If ws.Cells(i, "R").Value Like "*-*-*" Then
                            ws.Cells(i, "Q").Value = ws.Cells(i, "R").Value
                            ws.Cells(i, "R").ClearContents
                        End If
                        ' Location code
                        If ws.Cells(i, "U").Value <> "" Then
                            ws.Cells(i, "R").Value = ws.Cells(i, "U").Value
                            ws.Cells(i, "U").ClearContents
                        End If
                        ' Remove "Unnamed" columns
                        Dim col As Integer
                        For col = ws.Cells(1, Columns.count).End(xlToLeft).Column To 1 Step -1
                            If ws.Cells(1, col).Value Like "Unnamed*" Then
                                ws.Columns(col).Delete
                            End If
                        Next col
                    End If
                Next i
            End If

            ' === Categorize workbook ===
            If zipColFound Then zipFormattedList = zipFormattedList & vbCrLf & "- " & wb.Name
            If apexLogicApplied Then apexAppliedList = apexAppliedList & vbCrLf & "- " & wb.Name
            If arWestLogicApplied Then arWestAppliedList = arWestAppliedList & vbCrLf & "- " & wb.Name
            If Not zipColFound And Not apexLogicApplied And Not arWestLogicApplied Then
                untouchedList = untouchedList & vbCrLf & "- " & wb.Name
            End If
        End If
    Next wb

    ' === Summary Message ===
    Dim msg As String
    msg = "Processing Summary:" & vbCrLf & String(40, "-")

    If zipFormattedList <> "" Then
        msg = msg & vbCrLf & vbCrLf & "Workbooks with ZIP columns formatted:" & zipFormattedList
    Else
        msg = msg & vbCrLf & vbCrLf & "No workbooks had recognizable ZIP column headers."
    End If

    If apexAppliedList <> "" Then
        msg = msg & vbCrLf & vbCrLf & "Workbooks that received APEX processing:" & apexAppliedList
    End If

    If arWestAppliedList <> "" Then
        msg = msg & vbCrLf & vbCrLf & "Workbooks with arWestRestaurant corrections:" & arWestAppliedList
    End If

    If untouchedList <> "" Then
        msg = msg & vbCrLf & vbCrLf & "Workbooks with no changes:" & untouchedList
    End If

    MsgBox msg, vbInformation, "ZIP + APEX + arWest Processing Summary"
End Sub


