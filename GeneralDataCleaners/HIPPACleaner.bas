Attribute VB_Name = "HIPPACleaner"
Sub HIPPACleaner()
'
' ConsultReport_HIPPA_Edit Macro
' makes viable for sending without HIPPA violating
'

'
    Columns("F:F").Select
    Selection.Delete Shift:=xlToLeft
    Columns("G:M").Select
    Selection.Delete Shift:=xlToLeft
    Columns("A:A").EntireColumn.AutoFit
    Columns("B:B").EntireColumn.AutoFit
    Columns("C:C").EntireColumn.AutoFit
    Columns("E:E").EntireColumn.AutoFit
    Columns("D:D").EntireColumn.AutoFit
    Rows("1:1").Select
    With Selection.Interior
        .pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range("A1:G1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("$A$1:$G$37"), , xlYes).Name = _
        "Table2"
    Range("Table2[#All]").Select
    ActiveSheet.ListObjects("Table2").TableStyle = "TableStyleMedium1"
    Selection.AutoFilter
    Range("D21").Select
End Sub

