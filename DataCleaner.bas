Attribute VB_Name = "DataCleaner"
Sub DataCleaner()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    Dim headers As Range, cell As Range
    Dim lastRow As Long, col As Long
    Dim headerMap As Object
    Set headerMap = CreateObject("Scripting.Dictionary")

    ' Define target headers and flexible matching
    Dim key As Variant
    Dim headerLabels As Object
    Set headerLabels = CreateObject("Scripting.Dictionary")
    
    headerLabels.Add "Group ID", Array("Group ID", "Grp ID", "GrpID")
    headerLabels.Add "Product Code", Array("Product Code", "Prod Code", "ProductCode")
    headerLabels.Add "Active Date", Array("Active Date", "Start Date", "Activation Date")
    headerLabels.Add "Inactive Date", Array("Inactive Date", "End Date", "Deactivation Date")
    headerLabels.Add "First Name", Array("First Name", "FName", "Given Name")
    headerLabels.Add "Middle Name", Array("Middle Name", "MName", "Mid Name")
    headerLabels.Add "Last Name", Array("Last Name", "LName", "Surname")
    headerLabels.Add "Date of Birth", Array("Date of Birth", "DOB", "Birthdate")
    headerLabels.Add "Gender", Array("Gender", "Sex")
    headerLabels.Add "Address1", Array("Address1", "Address Line 1", "Addr1")
    headerLabels.Add "Address2", Array("Address2", "Address Line 2", "Addr2")
    headerLabels.Add "City", Array("City", "Town")
    headerLabels.Add "State", Array("State", "Province")
    headerLabels.Add "Zip", Array("Zip", "Zip Code", "Postal Code")
    headerLabels.Add "Phone", Array("Phone", "Phone Number", "Tel")

    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    Set headers = ws.Range("1:1")

    ' Identify column positions
    For Each cell In headers
        For Each key In headerLabels.Keys
            For Each Label In headerLabels(key)
                If LCase(Trim(cell.Value)) = LCase(Label) Then
                    headerMap(key) = cell.Column
                End If
            Next
        Next
    Next

    Dim i As Long, val As String, colIdx As Long
    Dim genderFilled As Boolean: genderFilled = False

    For Each key In headerMap.Keys
        colIdx = headerMap(key)
        For i = 2 To lastRow
            val = Trim(ws.Cells(i, colIdx).Value)
            val = Replace(val, ".", "")
            val = Replace(val, ",", "")
            val = Replace(val, "-", "")
            val = Replace(val, "(", "")
            val = Replace(val, ")", "")
            val = Replace(val, "/", "")
            val = Application.WorksheetFunction.Trim(val)
            
            Select Case key
                Case "Group ID"
                    If Len(val) = 6 And IsNumeric(val) Then
                        ws.Cells(i, colIdx).Value = val
                    Else
                        Debug.Print "Invalid Group ID at row " & i
                    End If
                Case "Product Code"
                    If Len(val) = 5 And IsNumeric(val) Then
                        ws.Cells(i, colIdx).Value = val
                    Else
                        Debug.Print "Invalid Product Code at row " & i
                    End If
                Case "Active Date", "Inactive Date", "Date of Birth"
                    If IsDate(ws.Cells(i, colIdx).Value) Then
                        ws.Cells(i, colIdx).NumberFormat = "mm/dd/yyyy"
                    Else
                        Debug.Print key & " not a valid date at row " & i
                    End If
                Case "First Name", "Middle Name", "Last Name", "Address1", "Address2", "City"
                    ws.Cells(i, colIdx).Value = Application.WorksheetFunction.Proper(val)
                Case "Gender"
                    val = UCase(Left(val, 1))
                    If val = "" Then
                        ws.Cells(i, colIdx).Value = "M"
                        genderFilled = True
                    ElseIf val = "M" Or val = "F" Or val = "U" Then
                        ws.Cells(i, colIdx).Value = val
                    Else
                        Debug.Print "Invalid gender at row " & i
                    End If
                Case "State"
                    If Len(val) <> 2 Or Not val Like "[A-Z][A-Z]" Then
                        Debug.Print "Invalid State at row " & i
                    End If
                    ws.Cells(i, colIdx).Value = UCase(val)
                Case "Zip"
                    If IsNumeric(val) Then
                        If Len(val) > 5 Then val = Left(val, 5)
                        ws.Cells(i, colIdx).Value = val
                    Else
                        Debug.Print "Invalid Zip at row " & i
                    End If
                Case "Phone"
                    val = Replace(val, " ", "")
                    If Len(val) = 10 And IsNumeric(val) Then
                        ws.Cells(i, colIdx).Value = "(" & Mid(val, 1, 3) & ") " & Mid(val, 4, 3) & "-" & Mid(val, 7)
                    Else
                        Debug.Print "Invalid phone number at row " & i
                    End If
            End Select
        Next i
    Next key

    If genderFilled Then
        MsgBox "Some missing Gender values were filled in as 'M'"
    End If
    MsgBox "Formatting complete. Check Immediate Window (Ctrl+G) for warnings."
End Sub

