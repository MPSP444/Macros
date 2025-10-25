Attribute VB_Name = "Module1"
Sub Макрос1()
Attribute Макрос1.VB_ProcData.VB_Invoke_Func = "г\n14"
'
' Макрос1 Макрос
'
' Сочетание клавиш: Ctrl+г
'
    Application.Run "PERSONAL.XLSB!ShowHiddenRows"
    ActiveWindow.SmallScroll Down:=-3
    Application.Run "PERSONAL.XLSB!ClearData"
    Application.Run "PERSONAL.XLSB!HideRows"
End Sub

Sub Макрос3()
Attribute Макрос3.VB_ProcData.VB_Invoke_Func = "й\n14"
'
' Макрос3 Макрос
'
' Сочетание клавиш: Ctrl+й
'
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.349986266670736
        .PatternTintAndShade = 0
    End With
End Sub
Sub Макрос4()
Attribute Макрос4.VB_ProcData.VB_Invoke_Func = "ц\n14"
'
' Макрос4 Макрос
'
' Сочетание клавиш: Ctrl+ц
'
       With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0.599993896298105
        .PatternTintAndShade = 0
    End With
End Sub
Sub Макрос5()
Attribute Макрос5.VB_ProcData.VB_Invoke_Func = "у\n14"
'
' Макрос5 Макрос
'
' Сочетание клавиш: Ctrl+у
'
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 49407
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent4
        .TintAndShade = 0.599993896298105
        .PatternTintAndShade = 0
    End With
End Sub
Sub Макрос6()
Attribute Макрос6.VB_ProcData.VB_Invoke_Func = "к\n14"
'
' Макрос6 Макрос
'
' Сочетание клавиш: Ctrl+к
'
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection.Font
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
    End With
End Sub
Sub Макрос7()
Attribute Макрос7.VB_ProcData.VB_Invoke_Func = "е\n14"
'
' Макрос7 Макрос
'
' Сочетание клавиш: Ctrl+е
'
    With Selection.Font
        .Color = -4569402
        .TintAndShade = 0
    End With
End Sub
Sub Макрос8()
Attribute Макрос8.VB_ProcData.VB_Invoke_Func = "н\n14"
'
' Макрос8 Макрос
'
' Сочетание клавиш: Ctrl+н
'
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 15773696
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
End Sub

Sub Макрос10()
Attribute Макрос10.VB_ProcData.VB_Invoke_Func = "р\n14"
'
' Макрос10 Макрос
'
' Сочетание клавиш: Ctrl+р
'
    Selection.UnMerge
End Sub



Sub Макрос15()
Attribute Макрос15.VB_ProcData.VB_Invoke_Func = "п\n14"
'
' Макрос15 Макрос
'
' Сочетание клавиш: Ctrl+п
'
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
    Selection.Font.Bold = False
    Selection.Font.Bold = True
End Sub
