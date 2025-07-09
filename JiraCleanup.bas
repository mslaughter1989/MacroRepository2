Attribute VB_Name = "JiraCleanup"
Sub JiraCleanup()
Attribute JiraCleanup.VB_ProcData.VB_Invoke_Func = " \n14"
'
' JiraCleanup Macro
'

'
    Selection.Replace What:="Custom field (", Replacement:="", LookAt:=xlPart _
        , SearchOrder:=xlByColumns, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:=")", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByColumns, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Range("CM2").Select
    ActiveWindow.ScrollColumn = 68
    ActiveWindow.ScrollColumn = 67
    ActiveWindow.ScrollColumn = 64
    ActiveWindow.ScrollColumn = 61
    ActiveWindow.ScrollColumn = 58
    ActiveWindow.ScrollColumn = 54
    ActiveWindow.ScrollColumn = 43
    ActiveWindow.ScrollColumn = 39
    ActiveWindow.ScrollColumn = 37
    ActiveWindow.ScrollColumn = 34
    ActiveWindow.ScrollColumn = 30
    ActiveWindow.ScrollColumn = 27
    ActiveWindow.ScrollColumn = 25
    ActiveWindow.ScrollColumn = 22
    ActiveWindow.ScrollColumn = 18
    ActiveWindow.ScrollColumn = 16
    ActiveWindow.ScrollColumn = 13
    ActiveWindow.ScrollColumn = 10
    ActiveWindow.ScrollColumn = 7
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 1
    Columns("A:A").Select
    Selection.Delete Shift:=xlToLeft
    Columns("B:B").Select
    Selection.Delete Shift:=xlToLeft
    Rows("1:1").Select
    Selection.AutoFilter
    Columns("B:B").Select
    Selection.Delete Shift:=xlToLeft
    Columns("C:Q").Select
    Selection.Delete Shift:=xlToLeft
    Columns("C:Q").EntireColumn.AutoFit
    Columns("D:EU").Select
    Selection.Delete Shift:=xlToLeft
    Columns("F:DL").Select
    Selection.Delete Shift:=xlToLeft
    Range("E1").Select
    Columns("D:D").EntireColumn.AutoFit
    Columns("E:E").EntireColumn.AutoFit
End Sub

