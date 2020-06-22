Attribute VB_Name = "Module2"
Sub Macro1()
Attribute Macro1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro1 Macro
'

'
    Application.Left = 126.25
    Application.Top = -20
    Range("C15").Select
    Application.Left = 130.75
    Application.Top = 22
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 15773696
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    ActiveCell.FormulaR1C1 = "9"
    Range("C16").Select
    Application.Left = 115
    Application.Top = 22
End Sub
