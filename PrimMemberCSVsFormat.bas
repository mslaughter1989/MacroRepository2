Attribute VB_Name = "PrimMemberCSVsFormat"
Option Explicit

'-------------------------------------------------------------------
'  CleanMemberDataInPlace
'  • Keeps only required columns (accepts header variants)
'  • Filters MemberType <> "1"
'  • Filters GroupName to five approved values (ignores punctuation)
'  • Strips “T00:00:00.” from EffectiveStart and formats mm/dd/yyyy
'-------------------------------------------------------------------
Sub PrimMemberCSVsFormat()

    Dim ws              As Worksheet
    Dim hdrRow          As Range
    Dim i               As Long, lastRow As Long
    Dim hdrText         As String, norm As String
    Dim keepDict        As Object
    Dim memberTypeCol   As Long, effCol As Long, grpCol As Long
    Dim cell            As Range
    Dim groupDict       As Object
    Dim reg             As Object
    Dim v               As Variant          '<<< Variant for array loop
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    Set ws = ActiveSheet
    Set hdrRow = ws.Rows(1)
    
    '--- dictionary of headers to keep ---------------------------------------
    Set keepDict = CreateObject("Scripting.Dictionary")
    keepDict.CompareMode = vbTextCompare
    keepDict("MEMBERID") = True
    keepDict("FIRSTNAME") = True
    keepDict("LASTNAME") = True
    keepDict("MEMBERTYPE") = True
    keepDict("EFFECTIVESTART") = True
    keepDict("GROUPNAME") = True
    keepDict("ACTIVATIONDATE") = True
    keepDict("LASTACTIVE") = True
    keepDict("LOGINEVENTS") = True
    
    '--- remove unwanted columns (loop right-to-left) -------------------------
    For i = hdrRow.Columns.count To 1 Step -1
        hdrText = Trim$(hdrRow.Cells(1, i).Value)
        norm = Replace(Replace(UCase$(hdrText), " ", ""), "_", "")
        If Len(hdrText) = 0 Or Not keepDict.Exists(norm) Then ws.Columns(i).Delete
    Next i
    Set hdrRow = ws.Rows(1)           'refresh after deletions
    
    '--- locate key columns --------------------------------------------------
    For i = 1 To hdrRow.Columns.count
        hdrText = Trim$(hdrRow.Cells(1, i).Value)
        norm = Replace(Replace(UCase$(hdrText), " ", ""), "_", "")
        Select Case norm
            Case "MEMBERTYPE":     memberTypeCol = i
            Case "EFFECTIVESTART": effCol = i
            Case "GROUPNAME":      grpCol = i
        End Select
    Next i
    
    If memberTypeCol * effCol * grpCol = 0 Then
        MsgBox "One of MemberType, EffectiveStart, or GroupName headers is missing.", vbCritical
        GoTo CleanExit
    End If
    
    '--- 1) Filter MemberType so only “1” remains ----------------------------
    ws.AutoFilterMode = False
    hdrRow.AutoFilter Field:=memberTypeCol, Criteria1:="1"
    
    '--- 2) Filter GroupName to the approved list ----------------------------
    Set groupDict = CreateObject("Scripting.Dictionary")
    groupDict.CompareMode = vbTextCompare
    
    Set reg = CreateObject("VBScript.RegExp")
    reg.Global = True: reg.pattern = "[^A-Z0-9]"   'strip non-alphanumerics
    
    Dim approved As Variant
    approved = Array("Empire Paper", _
                     "RRM Design Group Inc", _
                     "Medallion Financial Corp", _
                     "WestCare Foundation, Inc", _
                     "Utah Physical Therapy")
    
    For Each v In approved                     '<<< Variant loop variable
        groupDict(UCase(reg.Replace(v, ""))) = True
    Next v
    
    lastRow = ws.Cells(ws.Rows.count, grpCol).End(xlUp).Row
    On Error Resume Next
    For Each cell In ws.Range(ws.Cells(2, grpCol), ws.Cells(lastRow, grpCol)) _
                        .SpecialCells(xlCellTypeVisible)
        norm = UCase(reg.Replace(CStr(cell.Value), ""))
        If Not groupDict.Exists(norm) Then cell.EntireRow.Hidden = True
    Next cell
    On Error GoTo 0
    
    '--- 3) Clean EffectiveStart ---------------------------------------------
    lastRow = ws.Cells(ws.Rows.count, effCol).End(xlUp).Row
    On Error Resume Next
    For Each cell In ws.Range(ws.Cells(2, effCol), ws.Cells(lastRow, effCol)) _
                        .SpecialCells(xlCellTypeVisible)
        If Len(cell.Value) >= 9 Then
            cell.Value = Left$(cell.Value, Len(cell.Value) - 9)
            If IsDate(cell.Value) Then cell.Value = CDate(cell.Value)
        End If
    Next cell
    On Error GoTo 0
    ws.Columns(effCol).NumberFormat = "mm/dd/yyyy"
    
    MsgBox "Done! Data cleaned and filtered in place.", vbInformation
    
CleanExit:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub


