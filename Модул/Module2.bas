Attribute VB_Name = "Module2"
Sub Макрос2()
Attribute Макрос2.VB_ProcData.VB_Invoke_Func = "ш\n14"
'
' Макрос2 Макрос
'
' Сочетание клавиш: Ctrl+ш
'
    Application.Run "PERSONAL.XLSB!ShowHiddenRows"
    Application.Run "PERSONAL.XLSB!ProcessSpecialRows"
    Application.Run "PERSONAL.XLSB!CalculateAllWarehouses.CalculateAllWarehouses"
    Application.Run "PERSONAL.XLSB!SortWarehouseData"
    ActiveWindow.SmallScroll Down:=-3
    Application.Run "PERSONAL.XLSB!HideRows"
End Sub
