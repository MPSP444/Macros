Attribute VB_Name = "ModuleUtilities"

' ===== ModuleUtilities =====
Option Explicit

Public Sub ShowHiddenRows()
    On Error Resume Next
    Application.ScreenUpdating = False
    
    With Worksheets("Ангар 5").Range("72:1000").EntireRow
        .Hidden = False
        .Parent.Parent.Windows(1).ScrollRow = 1000
    End With
    
    With Worksheets("Ангар 6").Range("74:1000").EntireRow
        .Hidden = False
        .Parent.Parent.Windows(1).ScrollRow = 74
    End With
    
    With Worksheets("Ангар 7").Range("60:1000").EntireRow
        .Hidden = False
        .Parent.Parent.Windows(1).ScrollRow = 60
    End With
    
    With Worksheets("Ангар 8").Range("59:1000").EntireRow
        .Hidden = False
        .Parent.Parent.Windows(1).ScrollRow = 59
    End With
    
    With Worksheets("Ангар 9").Range("83:1000").EntireRow
        .Hidden = False
        .Parent.Parent.Windows(1).ScrollRow = 83
    End With
    
    With Worksheets("Ангар 10").Range("82:1000").EntireRow
        .Hidden = False
        .Parent.Parent.Windows(1).ScrollRow = 82
    End With
    
    With Worksheets("Ангар 11").Range("69:1000").EntireRow
        .Hidden = False
        .Parent.Parent.Windows(1).ScrollRow = 69
    End With
    
    With Worksheets("Ангар 12").Range("70:1000").EntireRow
        .Hidden = False
        .Parent.Parent.Windows(1).ScrollRow = 70
    End With
    
    Application.ScreenUpdating = True
    MsgBox "Строки показаны!", vbInformation
End Sub

Public Sub HideRows()
    On Error Resume Next
    Application.ScreenUpdating = False
    
    ' Скрываем строки в каждом ангаре
    Worksheets("Ангар 5").rows("72:1000").Hidden = True
    Worksheets("Ангар 6").rows("74:1000").Hidden = True
    Worksheets("Ангар 7").rows("60:1000").Hidden = True
    Worksheets("Ангар 8").rows("59:1000").Hidden = True
    Worksheets("Ангар 9").rows("83:1000").Hidden = True
    Worksheets("Ангар 10").rows("82:1000").Hidden = True
    Worksheets("Ангар 11").rows("69:1000").Hidden = True
    Worksheets("Ангар 12").rows("70:1000").Hidden = True
    
    Application.ScreenUpdating = True
    MsgBox "Строки скрыты!", vbInformation
End Sub

Public Sub ClearData()
    On Error Resume Next
    Application.ScreenUpdating = False
    
    ' Очистка данных в каждом ангаре
    Worksheets("Ангар 5").Range("A72:B1000").value = ""
    Worksheets("Ангар 6").Range("A74:B1000").value = ""
    Worksheets("Ангар 7").Range("A60:B1000").value = ""
    Worksheets("Ангар 8").Range("A58:B1000").value = ""
    Worksheets("Ангар 9").Range("A83:B1000").value = ""
    Worksheets("Ангар 10").Range("A82:B1000").value = ""
    Worksheets("Ангар 11").Range("A69:B1000").value = ""
    Worksheets("Ангар 12").Range("A70:B1000").value = ""
    
    Application.ScreenUpdating = True
    MsgBox "Данные очищены!", vbInformation
End Sub
