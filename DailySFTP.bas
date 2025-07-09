Attribute VB_Name = "DailyFormatting"
Sub DailyFormatting()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim rngP As Range, cell As Range
    Dim dict As Object
    Dim i As Long, j As Long
    Dim lastRow As Long, lastCol As Long
    Dim headerCell As Range
    Dim cleanHeader As String
    Dim zipColFound As Boolean, apexLogicApplied As Boolean
    Dim keywordList As Variant
    Dim keyword As Variant

    Dim zipFormattedList As String
    Dim apexAppliedList As String
    Dim untouchedList As String

    ' Initialize tracking
    keywordList = Array("zip", "zipcode", "zip code", "postalcode", "postal code")
    zipFormattedList = ""
    apexAppliedList = ""
    untouchedList = ""

    For Each wb In Application.Workbooks
        If wb.Sheets.count > 0 Then
            Set ws = wb.Sheets(1)
            zipColFound = False
            apexLogicApplied = False

            ' === Format ZIP code columns based on header match ===
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

            ' === Run APEX logic if name contains 'APEX' ===
            If InStr(1, UCase(wb.Name), "APEX") > 0 Then
                apexLogicApplied = True

                ' Step 1: Count duplicates in Column P
                Set rngP = ws.Range("P2:P" & ws.Cells(ws.Rows.count, "P").End(xlUp).Row)
                Set dict = CreateObject("Scripting.Dictionary")

                For Each cell In rngP
                    If Not dict.Exists(cell.Value) Then
                        dict.Add cell.Value, 1
                    Else
                        dict(cell.Value) = dict(cell.Value) + 1
                    End If
                Next cell

                ' Step 2: Delete rows with duplicate P and data in N
                For i = ws.Cells(ws.Rows.count, "P").End(xlUp).Row To 2 Step -1
                    If dict.Exists(ws.Cells(i, "P").Value) And dict(ws.Cells(i, "P").Value) > 1 _
                        And ws.Cells(i, "N").Value <> "" Then
                        ws.Rows(i).Delete
                    End If
                Next i

                ' Step 3: For remaining duplicates in P, delete row with smaller M
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

            ' === Categorize workbook based on results ===
            If zipColFound Then
                zipFormattedList = zipFormattedList & vbCrLf & "- " & wb.Name
            End If

            If apexLogicApplied Then
                apexAppliedList = apexAppliedList & vbCrLf & "- " & wb.Name
            End If

            If Not zipColFound And Not apexLogicApplied Then
                untouchedList = untouchedList & vbCrLf & "- " & wb.Name
            End If
        End If
    Next wb

    ' === Build and show the summary ===
    Dim msg As String
    msg = "Processing Summary:" & vbCrLf & String(40, "-")

    If zipFormattedList <> "" Then
        msg = msg & vbCrLf & vbCrLf & "Workbooks with ZIP columns formatted:" & zipFormattedList
    Else
        msg = msg & vbCrLf & vbCrLf & "No workbooks had recognizable ZIP column headers."
    End If

    If apexAppliedList <> "" Then
        msg = msg & vbCrLf & vbCrLf & "Workbooks that received APEX processing:" & apexAppliedList
    Else
        msg = msg & vbCrLf & vbCrLf & "No workbooks required APEX processing."
    End If

    If untouchedList <> "" Then
        msg = msg & vbCrLf & vbCrLf & "Workbooks with no changes:" & untouchedList
    End If

    MsgBox msg, vbInformation, "ZIP + APEX Macro Results"
End Sub

