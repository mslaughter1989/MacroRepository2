Attribute VB_Name = "PeopleFirstEdit"
Option Explicit

'===========================  MAIN PROCEDURE  ===========================
Sub PeopleFirstEdit()

    'Accepts  PF1_RECURO_Eligibility_######.csv   or  PF1_RECURO_Eligibility_########.csv
    Const PATTERN_6 As String = "PF1_RECURO_Eligibility_######.csv*"
    Const PATTERN_8 As String = "PF1_RECURO_Eligibility_########.csv*"

    Dim wb  As Workbook, ws As Worksheet
    Dim compCol As Long, prodCol As Long, zipCol As Long
    Dim lastRow As Long, r As Long, key As String
    
    '--- mapping of cleaned-up company names ? desired product code
    Dim map As Object: Set map = CreateObject("Scripting.Dictionary")
    map(CleanStr("SOLIDCORE HOLDINGS LLC")) = "39658"
    map(CleanStr("GEORGETOWN HILL CHILD CARE CENTER INC")) = "33212"
    map(CleanStr("EASY ICE LLC")) = "33212"
    map(CleanStr("BOOMTOWN NETWORK INC")) = "33212"
    
    '----- loop through every OPEN workbook
    For Each wb In Application.Workbooks
        If (wb.Name Like PATTERN_6) Or (wb.Name Like PATTERN_8) Then
            
            Set ws = wb.Worksheets(1)           'first (only) sheet in the CSV
            
            'HEADER COLUMNS -------------------------------------------------
            compCol = HeaderCol(ws.Rows(1), "Company Name")
            prodCol = HeaderCol(ws.Rows(1), "Product Code")
            zipCol = HeaderCol(ws.Rows(1), "Zip Code")
            
            If compCol = 0 Or prodCol = 0 Then
                MsgBox "Couldn’t find the Company-Name or Product-Code column in " & wb.Name, vbExclamation
                GoTo NextWorkbook
            End If
            
            'DATA -----------------------------------------------------------
            lastRow = ws.Cells(ws.Rows.count, compCol).End(xlUp).Row
            
            For r = 2 To lastRow                       'skip header row
                key = CleanStr(ws.Cells(r, compCol).Value)
                If map.Exists(key) Then
                    ws.Cells(r, prodCol).Value = map(key)
                End If
            Next r
            
            'FORMAT ZIP-CODE COLUMN ----------------------------------------
            If zipCol > 0 Then ws.Columns(zipCol).NumberFormat = "00000"
            
            MsgBox "? Product codes updated in " & wb.Name, vbInformation
            
        End If
NextWorkbook:
    Next wb
    
End Sub

'===========================  HELPER ROUTINES  ===========================

'Returns the 1-based column number whose cleaned-up header matches targetHeader
Private Function HeaderCol(headerRow As Range, targetHeader As String) As Long
    Dim c As Range
    For Each c In headerRow.Cells
        If Len(c.Value) = 0 Then Exit For                 'stop after trailing blanks
        If CleanStr(c.Value) = CleanStr(targetHeader) Then
            HeaderCol = c.Column: Exit Function
        End If
    Next c
    HeaderCol = 0                                         'not found
End Function

'Normalises strings: uppercase, strip spaces, commas, tabs, non-breaking spaces & line-breaks
Private Function CleanStr(s As String) As String
    Dim t As String
    t = UCase$(s)
    t = Replace(t, ",", "")
    t = Replace(t, Chr(160), "")      ' non-breaking space
    t = Replace(t, " ", "")
    t = Replace(t, vbTab, "")
    t = Replace(t, vbCr, "")
    t = Replace(t, vbLf, "")
    CleanStr = Trim$(t)
End Function

