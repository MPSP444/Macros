Attribute VB_Name = "Module12"
Sub Макрос14()
Attribute Макрос14.VB_ProcData.VB_Invoke_Func = "а\n14"
'
' Макрос14 Макрос
'
' Сочетание клавиш: Ctrl+а
'
    Application.Run "PERSONAL.XLSB!ЗапуститьОбработкуПоАнгарам"
    Range("C28").Select
    Application.Run "PERSONAL.XLSB!ЗапуститьОбработкуПоАнгарам"
    Range("C29").Select
End Sub
