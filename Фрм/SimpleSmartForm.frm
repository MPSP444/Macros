VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SimpleSmartForm 
   Caption         =   "UserForm4"
   ClientHeight    =   5760
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11520
   OleObjectBlob   =   "SimpleSmartForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SimpleSmartForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ===== SimpleSmartForm - ������ ������ � ������� ������ =====
Option Explicit

Private lastReport As String  ' ��� �������� ���������� ������

' ������ ����������������� ���� ���������� ������� �����
' �������: 0=ProductName, 1=Batch, 2=Quantity, 3=Location

' ===== ����� ���������� ������ ������ =====
Private Sub CommandButton10_Click()
    If lastReport = "" Then
        MsgBox "��� ������ ��� ������!", vbInformation
        Exit Sub
    End If
    
    ' �������� �������� �����
    PrintBeautifulReport lastReport
End Sub

Private Sub CommandButton9_Click()
    If lastReport = "" Then
        MsgBox "��� ������ ��� ������!", vbInformation
        Exit Sub
    End If
    
    ' ���������� ������ �����
    ShowReportForm lastReport

End Sub

Private Sub ShowReportForm(reportText As String)
    On Error GoTo ErrorHandler
    
    ' ������������� ���� ��� ����������
    Dim filePath As String
    filePath = "C:\Users\mukam\Desktop\�������_��_��\����������.txt"
    
    ' ��������� ������������� �����
    Dim folderPath As String
    folderPath = "C:\Users\mukam\Desktop\�������_��_��"
    
    If Dir(folderPath, vbDirectory) = "" Then
        MkDir folderPath
    End If
    
    ' ��������� ���� � ����� � ������
    Dim fullReport As String
    fullReport = "=== ����� � ���������� ������� ===" & vbNewLine
    fullReport = fullReport & "����: " & Format(Now, "dd.mm.yyyy hh:mm:ss") & vbNewLine
    fullReport = fullReport & String(50, "=") & vbNewLine & vbNewLine
    fullReport = fullReport & reportText
    
    ' ��������� � ����
    Dim fileNum As Integer
    fileNum = FreeFile
    Open filePath For Output As #fileNum
    Print #fileNum, fullReport
    Close #fileNum
    
    ' ��������� ���� � ��������
    Shell "notepad.exe " & filePath, vbNormalFocus
    
    MsgBox "����� ��������:" & vbNewLine & filePath & vbNewLine & vbNewLine & _
           "���� ������ � �������� - ������ ����������� ������ �����", vbInformation, "����� ��������"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "������ ���������� ������:" & vbNewLine & Err.description & vbNewLine & vbNewLine & _
           "��������� ����: " & filePath, vbExclamation, "������"
End Sub

Private Sub UserForm_Initialize()
    ' === ������ ����� (����������� ��� ����� ������) ===
    Me.Width = 600
    Me.Height = 680  ' ��������� ��� ����� ������
    Me.Caption = "����� ���������� �������"
    
    ' === 1. ���������� (Label1) ===
    With Label1
        .Caption = "������� ������ ��� �������� ����� (������: ����� - ������ - ����������):"
        .Top = 10
        .Left = 10
        .Width = 570
        .Height = 30
        .Font.Size = 12
        .Font.Bold = True
        .WordWrap = True
    End With
    
    ' === 2. ��������� ���� ��� ����� (TextBox1) ===
    With TextBox1
        .MultiLine = True
        .ScrollBars = fmScrollBarsBoth
        .Top = 50
        .Left = 10
        .Width = 570
        .Height = 120
        .Font.Name = "Consolas"
        .Font.Size = 11
        ' ������ ��� ������������
        .text = "��������� - ���03 - 720 (720)" & vbNewLine & _
"��������� - ���02 - 690 (720)" & vbNewLine & _
"��������� - ���03 - 400 (720)" & vbNewLine & _
"�������� - ���01 - 720 (720)" & vbNewLine & _
"�������� - ���02 - 110 (720)" & vbNewLine & _
"��������� ��������� - ���13 - 720 (720)" & vbNewLine & _
"��������� ��������� - ���13 - 720 (720)" & vbNewLine & _
"��������� ��������� - ���13 - 720 (720)" & vbNewLine & _
"��������� ��������� - ���27 - 720 (720)" & vbNewLine & _
"��������� ��������� - ���27 - 720 (720)" & vbNewLine & _
"��������� ��������� - ���27 - 720 (720)"
    End With
    
    ' === 3. ��������� ������� (Label2) ===
    With Label2
        .Caption = "?? �������� ������ ��� ����������:"
        .Top = 180
        .Left = 10
        .Width = 570
        .Height = 20
        .Font.Size = 11
        .Font.Bold = True
        .ForeColor = RGB(0, 100, 0)
    End With
    
    ' === 4. ������� ������� (CheckBox1-8) ===
    ' ������ ���: ������ 5,6,7,8
    With CheckBox1  ' ����� 5
        .Caption = "����� 5"
        .Top = 210
        .Left = 15
        .Width = 80
        .Height = 20
        .value = True
    End With
    
    With CheckBox2  ' ����� 6
        .Caption = "����� 6"
        .Top = 210
        .Left = 105
        .Width = 80
        .Height = 20
        .value = True
    End With
    
    With CheckBox3  ' ����� 7
        .Caption = "����� 7"
        .Top = 210
        .Left = 195
        .Width = 80
        .Height = 20
        .value = True
    End With
    
    With CheckBox4  ' ����� 8
        .Caption = "����� 8"
        .Top = 210
        .Left = 285
        .Width = 80
        .Height = 20
        .value = True
    End With
    
    ' ������ ���: ������ 9,10,11,12
    With CheckBox5  ' ����� 9
        .Caption = "����� 9"
        .Top = 240
        .Left = 15
        .Width = 80
        .Height = 20
        .value = True
    End With
    
    With CheckBox6  ' ����� 10
        .Caption = "����� 10"
        .Top = 240
        .Left = 105
        .Width = 80
        .Height = 20
        .value = True
    End With
    
    With CheckBox7  ' ����� 11
        .Caption = "����� 11"
        .Top = 240
        .Left = 195
        .Width = 80
        .Height = 20
        .value = True
    End With
    
    With CheckBox8  ' ����� 12
        .Caption = "����� 12"
        .Top = 240
        .Left = 285
        .Width = 80
        .Height = 20
        .value = True
    End With
    
    ' === 5. ������ ���������� �������� ===
    With CommandButton4  ' ���
        .Caption = "���"
        .Top = 240
        .Left = 380
        .Width = 50
        .Height = 20
        .BackColor = RGB(200, 255, 200)
    End With
    
    With CommandButton5  ' ���
        .Caption = "���"
        .Top = 240
        .Left = 440
        .Width = 50
        .Height = 20
        .BackColor = RGB(255, 200, 200)
    End With
    
    ' === 6. ������� ������ ===
    With CommandButton1  ' ����������
        .Caption = "?? ����������"
        .Top = 280
        .Left = 10
        .Width = 150
        .Height = 40
        .Font.Bold = True
        .Font.Size = 12
        .BackColor = RGB(0, 200, 0)
        .ForeColor = RGB(255, 255, 255)
    End With
    
    With CommandButton2  ' ��������
        .Caption = "??? ��������"
        .Top = 280
        .Left = 170
        .Width = 150
        .Height = 40
        .Font.Bold = True
        .Font.Size = 12
        .BackColor = RGB(255, 150, 0)
        .ForeColor = RGB(255, 255, 255)
    End With
    
    With CommandButton3  ' �����
        .Caption = "?? �����"
        .Top = 280
        .Left = 330
        .Width = 150
        .Height = 40
        .Font.Bold = True
        .Font.Size = 12
        .BackColor = RGB(0, 150, 255)
        .ForeColor = RGB(255, 255, 255)
        .Enabled = False
    End With
    
    ' === 7. ������ (Label3) ===
    With Label3
        .Caption = "����� � ������. �������� ������ � ������� ������."
        .Top = 340
        .Left = 10
        .Width = 570
        .Height = 60
        .Font.Size = 10
        .BackColor = RGB(240, 240, 240)
        .BorderStyle = fmBorderStyleSingle
        .TextAlign = fmTextAlignCenter
    End With
    
    ' === 8. �������������� ������ ===
    With CommandButton6  ' ���������
        .Caption = "?? ���������"
        .Top = 420
        .Left = 10
        .Width = 100
        .Height = 30
    End With
    
    With CommandButton7  ' ������
        .Caption = "? ������"
        .Top = 420
        .Left = 120
        .Width = 100
        .Height = 30
    End With
    
    ' === ����� ������ ������ ===
    With CommandButton10  ' ������
        .Caption = "??? ������"
        .Top = 420
        .Left = 230
        .Width = 100
        .Height = 30
        .Font.Bold = True
        .BackColor = RGB(100, 149, 237)  ' ����� ����
        .ForeColor = RGB(255, 255, 255)
        .Enabled = False  ' ���������� ����� ����������
    End With
    
    With CommandButton9  ' ���������
        .Caption = "?? ���������"
        .Top = 420
        .Left = 340
        .Width = 100
        .Height = 30
        .Font.Bold = True
        .BackColor = RGB(255, 165, 0)  ' ��������� ����
        .ForeColor = RGB(255, 255, 255)
        .Enabled = False  ' ���������� ����� ����������
    End With
    
    With CommandButton8  ' �������
        .Caption = "? �������"
        .Top = 420
        .Left = 450
        .Width = 100
        .Height = 30
    End With
    
    ' === ������������� ===
    Call LoadWarehouseSettings
    Call UpdateWarehouseStatus
End Sub

' ===== �������� ����������� =====

Private Sub CommandButton1_Click()  ' ����������
    ' �������� ������
    If Trim(TextBox1.text) = "" Then
        MsgBox "������� ������ ��� ����������!", vbExclamation
        TextBox1.SetFocus
        Exit Sub
    End If
    
    ' �������� Excel
    If ActiveWorkbook Is Nothing Then
        MsgBox "�������� ���� � ��������!", vbExclamation
        Exit Sub
    End If
    
    ' �������� �������� �������
    If GetActiveWarehouseCount() = 0 Then
        MsgBox "?? �� ������ �� ���� ����� ��� ����������!" & vbNewLine & _
               "�������� ���� �� ���� �����.", vbExclamation
        Exit Sub
    End If
    
    ' ������������� � ����������� �� �������� �������
    Dim activeList As String
    activeList = GetActiveWarehousesList()
    
    If MsgBox("���������� ������ � �������� �������?" & vbNewLine & vbNewLine & _
              "�������� ������: " & activeList, vbYesNo + vbQuestion) = vbNo Then
        Exit Sub
    End If
    
    ' ��������� ������
    Label3.Caption = "? ����������� ���������� � " & GetActiveWarehouseCount() & " �������..."
    Label3.BackColor = RGB(255, 255, 200)
    DoEvents
    
    ' ������������� ������
    Call ModuleProductInfo.InitializeProductDatabase
    Call ModuleWarehouseCapacity.InitializeWarehouseCapacities
    
    ' �������������� ��������� �������
    Call SyncWarehouseStates
    
    ' ��������� ����������
    On Error GoTo ErrorHandler
    lastReport = ModuleSmartPlacement.SmartPlaceProducts(TextBox1.text)
    
    ' ���������� ���������
    If InStr(lastReport, "�������") > 0 Then
        Label3.Caption = "? ���������� ���������! ������� ����� ��� �������."
        Label3.BackColor = RGB(200, 255, 200)
        CommandButton3.Enabled = True
        CommandButton9.Enabled = True   ' �������� ������ ���������
        CommandButton10.Enabled = True  ' �������� ������ ������
        
        ' ���������� ������� ���������
        Dim successCount As Integer
        successCount = CountSuccessful(lastReport)
        MsgBox "������� ���������: " & successCount & " �������" & vbNewLine & _
               "� �������: " & activeList & vbNewLine & _
               "������ �������� ������ ������", vbInformation
    Else
        Label3.Caption = "? ������ ����������"
        Label3.BackColor = RGB(255, 200, 200)
        MsgBox lastReport, vbExclamation
    End If
    
    Exit Sub
    
ErrorHandler:
    Label3.Caption = "? ������: " & Err.description
    Label3.BackColor = RGB(255, 200, 200)
    MsgBox "������: " & Err.description, vbCritical
End Sub

Private Sub CommandButton2_Click()  ' ��������
    TextBox1.text = ""
    Label3.Caption = "���� �������. ����� � ����� �������."
    Label3.BackColor = RGB(240, 240, 240)
    CommandButton3.Enabled = False
    CommandButton9.Enabled = False   ' ��������� ������ ���������
    CommandButton10.Enabled = False  ' ��������� ������ ������
    lastReport = ""
    TextBox1.SetFocus
End Sub

Private Sub CommandButton3_Click()  ' �����
    If lastReport = "" Then
        MsgBox "��� ������ ��� ������!", vbInformation
        Exit Sub
    End If
    
    ' ���������� �����
    MsgBox lastReport, vbInformation, "��������� ����� � ����������"
End Sub

Private Sub CommandButton4_Click()  ' ���
    CheckBox1.value = True   ' ����� 5
    CheckBox2.value = True   ' ����� 6
    CheckBox3.value = True   ' ����� 7
    CheckBox4.value = True   ' ����� 8
    CheckBox5.value = True   ' ����� 9
    CheckBox6.value = True   ' ����� 10
    CheckBox7.value = True   ' ����� 11
    CheckBox8.value = True   ' ����� 12
    Call UpdateWarehouseStatus
End Sub

Private Sub CommandButton5_Click()  ' ���
    CheckBox1.value = False  ' ����� 5
    CheckBox2.value = False  ' ����� 6
    CheckBox3.value = False  ' ����� 7
    CheckBox4.value = False  ' ����� 8
    CheckBox5.value = False  ' ����� 9
    CheckBox6.value = False  ' ����� 10
    CheckBox7.value = False  ' ����� 11
    CheckBox8.value = False  ' ����� 12
    Call UpdateWarehouseStatus
End Sub

Private Sub CommandButton6_Click()  ' ���������
    Dim info As String
    info = "?? ��������� �������" & vbNewLine & vbNewLine
    info = info & "�������� ������: " & GetActiveWarehousesList() & vbNewLine
    info = info & "����� ����������: " & GetActiveWarehouseCount() & " �� 8" & vbNewLine & vbNewLine
    info = info & "?? ��������� ������������� �����������" & vbNewLine
    info = info & "?? ��� ��������� ������� ����������� ModuleWarehouseCapacity"
    
    MsgBox info, vbInformation, "���������� � ����������"
End Sub

Private Sub CommandButton7_Click()  ' ������
    Dim help As String
    help = "? ������� �� ������ ����������" & vbNewLine & vbNewLine
    help = help & "?? ������ �����:" & vbNewLine
    help = help & "����� - ������ - ����������" & vbNewLine & vbNewLine
    help = help & "?? �������:" & vbNewLine
    help = help & "�������� - ���13 - 720" & vbNewLine
    help = help & "������ - ���07 - 500" & vbNewLine
    help = help & "�������� - ���25 - 360" & vbNewLine & vbNewLine
    help = help & "?? �����������:" & vbNewLine
    help = help & "� ������� ���� ������ ����������� �����" & vbNewLine
    help = help & "� ��������� �������������� ������ �������" & vbNewLine
    help = help & "� ��������� ������ � ��������� �������" & vbNewLine
    help = help & "� �������� ����� ������ ������ ������" & vbNewLine & vbNewLine
    help = help & "??? ������:" & vbNewLine
    help = help & "� ������ '������' ������� �������� ��������� �����" & vbNewLine
    help = help & "� �������������� ���������� �� ������������" & vbNewLine
    help = help & "� ��������� ����������� �� ���������" & vbNewLine & vbNewLine
    help = help & "?? ���������:" & vbNewLine
    help = help & "� ��������� ����� � ��������� ����" & vbNewLine
    help = help & "� ������������� ��������� � ��������"
    
    MsgBox help, vbInformation, "������"
End Sub

Private Sub CommandButton8_Click()  ' �������
    If MsgBox("������� ����� ������ ����������?", vbYesNo + vbQuestion) = vbYes Then
        Call SaveWarehouseSettings
        Unload Me
    End If
End Sub

' ===== ����������� ������� =====

Private Sub CheckBox1_Click()  ' ����� 5
    Call UpdateWarehouseStatus
    Call ModuleWarehouseCapacity.SetWarehouseActive("5", CheckBox1.value)
End Sub

Private Sub CheckBox2_Click()  ' ����� 6
    Call UpdateWarehouseStatus
    Call ModuleWarehouseCapacity.SetWarehouseActive("6", CheckBox2.value)
End Sub

Private Sub CheckBox3_Click()  ' ����� 7
    Call UpdateWarehouseStatus
    Call ModuleWarehouseCapacity.SetWarehouseActive("7", CheckBox3.value)
End Sub

Private Sub CheckBox4_Click()  ' ����� 8
    Call UpdateWarehouseStatus
    Call ModuleWarehouseCapacity.SetWarehouseActive("8", CheckBox4.value)
End Sub

Private Sub CheckBox5_Click()  ' ����� 9
    Call UpdateWarehouseStatus
    Call ModuleWarehouseCapacity.SetWarehouseActive("9", CheckBox5.value)
End Sub

Private Sub CheckBox6_Click()  ' ����� 10
    Call UpdateWarehouseStatus
    Call ModuleWarehouseCapacity.SetWarehouseActive("10", CheckBox6.value)
End Sub

Private Sub CheckBox7_Click()  ' ����� 11
    Call UpdateWarehouseStatus
    Call ModuleWarehouseCapacity.SetWarehouseActive("11", CheckBox7.value)
End Sub

Private Sub CheckBox8_Click()  ' ����� 12
    Call UpdateWarehouseStatus
    Call ModuleWarehouseCapacity.SetWarehouseActive("12", CheckBox8.value)
End Sub

' ===== ������� ��� �������� ������ =====

' ===== ������� �������� ������ =====
Private Sub PrintBeautifulReport(reportText As String)
    On Error GoTo ErrorHandler
    
    ' ������ ������ �� ������
    Dim placements As Collection
    Set placements = ParseReportData(reportText)
    
    If placements.count = 0 Then
        MsgBox "��� ������ ��� ������!", vbInformation
        Exit Sub
    End If
    
    ' ������� ����� ���� ��� ������
    Dim printSheet As Worksheet
    Set printSheet = CreatePrintSheet()
    
    If printSheet Is Nothing Then
        Exit Sub
    End If
    
    ' ��������� ������
    Call FillPrintSheet(printSheet, placements)
    
    ' ����������� �������
    Call FormatPrintSheet(printSheet, placements.count)
    
    ' ���������� ��������������� �������� � ������
    printSheet.Select
    
    MsgBox "?? ����� ����� � ������!" & vbNewLine & vbNewLine & _
           "� ������ ������������� �� ������������" & vbNewLine & _
           "� �������� ��������� ����������" & vbNewLine & _
           "� ������� ��� �����������" & vbNewLine & _
           "� ����� �������: " & placements.count & vbNewLine & vbNewLine & _
           "������� OK ��� ������...", vbInformation, "����� � ������"
    
    ' �������� ������ ������
    Application.Dialogs(xlDialogPrint).Show
    
    Exit Sub
    
ErrorHandler:
    MsgBox "������ �������� ������ ��� ������:" & vbNewLine & Err.description, vbCritical
End Sub

' ===== ������� ������ �� ������ =====
Private Function ParseReportData(reportText As String) As Collection
    Set ParseReportData = New Collection
    
    On Error GoTo ErrorHandler
    
    ' ��������� ����� �� ������
    Dim lines() As String
    lines = Split(reportText, vbNewLine)
    
    ' ���� ������ � ������������
    Dim foundPlacements As Boolean
    foundPlacements = False
    
    Dim i As Integer
    For i = 0 To UBound(lines)
        Dim line As String
        line = Trim(lines(i))
        
        ' ���� ������ ������ ����������
        If InStr(line, "���������� ������-��-�������:") > 0 Then
            foundPlacements = True
            GoTo NextLine
        End If
        
        ' ���� ����� ���������� � ������ �������� ������
        If foundPlacements And line <> "" And InStr(line, "---------") = 0 Then
            ' ���������, �� ����������� �� ����������
            If InStr(line, "������:") > 0 Or InStr(line, "�����������:") > 0 Then
                Exit For
            End If
            
            ' ������ ������ ����������
            Dim placement(3) As String  ' 0=ProductName, 1=Batch, 2=Quantity, 3=Location
            If ParsePlacementLine(line, placement) Then
                ParseReportData.Add placement
            End If
        End If
        
NextLine:
    Next i
    
    ' ��������� �� ������������
    If ParseReportData.count > 1 Then
        Set ParseReportData = SortPlacementsByName(ParseReportData)
    End If
    
    Exit Function
    
ErrorHandler:
    MsgBox "������ �������� ������: " & Err.description, vbCritical
    Set ParseReportData = New Collection
End Function

' ===== ������� ����� ������ ���������� =====
Private Function ParsePlacementLine(line As String, ByRef placement() As String) As Boolean
    On Error GoTo ErrorHandler
    
    ' ������ ������: "����� 6-28-���-1 - �������� (���08) - 110 "
    
    ' ������� "����� " � ������
    If InStr(line, "����� ") = 1 Then
        line = Mid(line, 7) ' ������� "����� "
    End If
    
    ' ��������� �� " - "
    Dim parts() As String
    parts = Split(line, " - ")
    
    If UBound(parts) < 2 Then
        ParsePlacementLine = False
        Exit Function
    End If
    
    ' �������� ����� ����������
    placement(3) = Trim(parts(0))  ' Location
    
    ' ������ ������� � ������: "�������� (���08)"
    Dim productPart As String
    productPart = Trim(parts(1))
    
    ' ���� ������ � �������
    Dim openBracket As Integer, closeBracket As Integer
    openBracket = InStrRev(productPart, "(")
    closeBracket = InStrRev(productPart, ")")
    
    If openBracket > 0 And closeBracket > openBracket Then
        placement(0) = Trim(Left(productPart, openBracket - 1))  ' ProductName
        placement(1) = Trim(Mid(productPart, openBracket + 1, closeBracket - openBracket - 1))  ' Batch
    Else
        placement(0) = productPart  ' ProductName
        placement(1) = ""  ' Batch
    End If
    
    ' �������� ����������
    placement(2) = Trim(parts(2))  ' Quantity
    
    ParsePlacementLine = True
    Exit Function
    
ErrorHandler:
    ParsePlacementLine = False
End Function

' ===== ���������� �� ������������ =====
Private Function SortPlacementsByName(unsortedCollection As Collection) As Collection
    On Error GoTo ErrorHandler
    
    ' ������� ������ ��� ����������
    Dim tempArray() As Variant
    ReDim tempArray(1 To unsortedCollection.count)
    
    ' �������� � ������
    Dim i As Integer
    For i = 1 To unsortedCollection.count
        tempArray(i) = unsortedCollection(i)
    Next i
    
    ' ===== ������������� ���������� =====
    Dim j As Integer
    Dim temp As Variant
    For i = 1 To UBound(tempArray) - 1
        For j = i + 1 To UBound(tempArray)
            Dim item1() As String, item2() As String
            item1 = tempArray(i)
            item2 = tempArray(j)

            ' ��������� ������ ������� �� Location
            Dim warehouse1 As Integer, warehouse2 As Integer
            warehouse1 = GetWarehouseFromLocation(item1(3))
            warehouse2 = GetWarehouseFromLocation(item2(3))

            ' ��������� ������ �����
            Dim row1 As Integer, row2 As Integer
            row1 = GetRowFromLocation(item1(3))
            row2 = GetRowFromLocation(item2(3))

            ' ��������� ����������
            Dim qty1 As Double, qty2 As Double
            qty1 = Val(item1(2))
            qty2 = Val(item2(2))

            Dim needSwap As Boolean
            needSwap = False

            ' ������� 1: ���������� �� �������
            If warehouse1 > warehouse2 Then
                needSwap = True
            ElseIf warehouse1 = warehouse2 Then
                ' ������� 2: ���� ������ ����������, ��������� �� ������ ���� (��������!)
                If row1 > row2 Then
                    needSwap = True
                ElseIf row1 = row2 Then
                    ' ������� 3: ���� ��� ����������, ��������� �� �������� ������
                    If UCase(item1(0)) > UCase(item2(0)) Then
                        needSwap = True
                    ElseIf UCase(item1(0)) = UCase(item2(0)) Then
                        ' ������� 4: ���� �������� ����������, ��������� �� ���������� (��������)
                        If qty1 < qty2 Then  ' ������� ���������� ���� ����� ��������
                            needSwap = True
                        End If
                    End If
                End If
            End If

            ' ������ ������� ���� �����
            If needSwap Then
                temp = tempArray(i)
                tempArray(i) = tempArray(j)
                tempArray(j) = temp
            End If
        Next j
    Next i
    
    ' ������� ����� ���������
    Set SortPlacementsByName = New Collection
    For i = 1 To UBound(tempArray)
        SortPlacementsByName.Add tempArray(i)
    Next i
    
    Exit Function
    
ErrorHandler:
    Set SortPlacementsByName = unsortedCollection
End Function
Private Function GetWarehouseFromLocation(location As String) As Integer
    On Error Resume Next

    ' �� ������ "6-38-�-1" ��������� "6"
    Dim parts() As String
    parts = Split(location, "-")

    If UBound(parts) >= 0 Then
        GetWarehouseFromLocation = CInt(parts(0))
    Else
        GetWarehouseFromLocation = 999  ' ���� ������ - �������� � �����
    End If

    On Error GoTo 0
End Function

' ===== Новая функция для извлечения номера ряда =====
Private Function GetRowFromLocation(location As String) As Integer
    On Error Resume Next

    ' �� ������ "6-38-�-1" ��������� "38"
    Dim parts() As String
    parts = Split(location, "-")

    If UBound(parts) >= 1 Then
        GetRowFromLocation = CInt(parts(1))
    Else
        GetRowFromLocation = 999  ' ���� ������ - �������� � �����
    End If

    On Error GoTo 0
End Function

' ===== �������� ����� ��� ������ =====
Private Function CreatePrintSheet() As Worksheet
    On Error GoTo ErrorHandler
    
    ' ������� ������ ���� ���� ����
    Dim sheetName As String
    sheetName = "�����_������"
    
    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
        If ws.Name = sheetName Then
            Application.DisplayAlerts = False
            ws.Delete
            Application.DisplayAlerts = True
            Exit For
        End If
    Next ws
    
    ' ������� ����� ����
    Set CreatePrintSheet = ActiveWorkbook.Worksheets.Add
    CreatePrintSheet.Name = sheetName
    
    ' ��������� ������
    With CreatePrintSheet.PageSetup
        .Orientation = xlPortrait ' ������� ����������
        .PaperSize = xlPaperA4
        .LeftMargin = Application.InchesToPoints(0.5)
        .RightMargin = Application.InchesToPoints(0.5)
        .TopMargin = Application.InchesToPoints(0.75)
        .BottomMargin = Application.InchesToPoints(0.75)
        .HeaderMargin = Application.InchesToPoints(0.3)
        .FooterMargin = Application.InchesToPoints(0.3)
        .PrintHeadings = False
        .PrintGridlines = False
        .CenterHorizontally = True
        .CenterVertically = True
    End With
    
    Exit Function
    
ErrorHandler:
    MsgBox "������ �������� �����: " & Err.description, vbCritical
    Set CreatePrintSheet = Nothing
End Function

' ===== ���������� ����� ������� =====
Private Sub FillPrintSheet(printSheet As Worksheet, placements As Collection)
    On Error GoTo ErrorHandler
    
    With printSheet
        ' ������� ����
        .Cells.Clear
        
        ' === ��������� ������ ===
        .Cells(1, 1).value = "����� � ���������� �������"
        .Cells(2, 1).value = "����: " & Format(Now, "dd.mm.yyyy hh:mm:ss")
        .Cells(3, 1).value = "�������� ������: " & GetActiveWarehousesList()
        
        ' === ��������� ������� ===
        Dim headerRow As Long
        headerRow = 6
        
        .Cells(headerRow, 1).value = "������������"
        .Cells(headerRow, 2).value = "������"
        .Cells(headerRow, 3).value = "���-��"
        .Cells(headerRow, 4).value = "����"
        .Cells(headerRow, 5).value = "�����"
        
        ' === ������ ===
        Dim dataRow As Long
        dataRow = headerRow + 1
        
        Dim i As Integer
        For i = 1 To placements.count
            Dim placement() As String
            placement = placements(i)
            
            .Cells(dataRow, 1).value = placement(0)  ' ProductName
            .Cells(dataRow, 2).value = "(" & placement(1) & ")"  ' Batch
            .Cells(dataRow, 3).value = placement(2)  ' Quantity
            .Cells(dataRow, 4).value = ">"  ' �������
            .Cells(dataRow, 5).value = placement(3)  ' Location
            
            dataRow = dataRow + 1
        Next i
        
    End With
    
    Exit Sub
    
ErrorHandler:
    MsgBox "������ ���������� ������: " & Err.description, vbCritical
End Sub

' ===== �������������� ������� =====
Private Sub FormatPrintSheet(printSheet As Worksheet, dataCount As Long)
    On Error GoTo ErrorHandler
    
    With printSheet
        ' === ��������� ������ ===
        With .Range("A1:E1")
            .Merge
            .Font.Size = 16
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
            .Interior.Color = RGB(0, 100, 200)
            .Font.Color = RGB(255, 255, 255)
            .RowHeight = 30
        End With
        
        With .Range("A2:E2")
            .Merge
            .Font.Size = 12
            .HorizontalAlignment = xlCenter
            .Interior.Color = RGB(220, 230, 241)
            .RowHeight = 20
        End With
        
        With .Range("A3:E3")
            .Merge
            .Font.Size = 12
            .HorizontalAlignment = xlCenter
            .Interior.Color = RGB(220, 230, 241)
            .RowHeight = 20
        End With
        
        ' === ��������� ������� ===
        Dim headerRange As String
        headerRange = "A6:E6"
        
        With .Range(headerRange)
            .Font.Bold = True
            .Font.Size = 12
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Interior.Color = RGB(79, 129, 189)
            .Font.Color = RGB(255, 255, 255)
            .RowHeight = 30
        End With
        
        ' === ������ ������� ===
        If dataCount > 0 Then
            Dim dataRange As String
            dataRange = "A7:E" & (6 + dataCount)
            
            With .Range(dataRange)
                .Font.Size = 11
                .VerticalAlignment = xlCenter
                .RowHeight = 25
            End With
            
            ' ������������ �������
            .columns("A").HorizontalAlignment = xlLeft     ' ������������ - �����
            .columns("B").HorizontalAlignment = xlCenter   ' ������ - �� ������
            .columns("C").HorizontalAlignment = xlCenter   ' ���������� - �� ������
            .columns("D").HorizontalAlignment = xlCenter   ' ������� - �� ������
            .columns("E").HorizontalAlignment = xlCenter   ' ����� - �� ������
            
            ' ���� �������
            .columns("D").Font.Color = RGB(255, 0, 0)
            .columns("D").Font.Size = 14
            .columns("D").Font.Bold = True
            
            ' ������������ ����� �����
            Dim row As Long
            For row = 7 To 6 + dataCount
                If row Mod 2 = 0 Then
                    .Range("A" & row & ":E" & row).Interior.Color = RGB(242, 242, 242)
                End If
            Next row
        End If
        
        ' === ������� ������� ===
        Dim tableRange As String
        tableRange = "A6:E" & (6 + dataCount)
        
        With .Range(tableRange).Borders
            .LineStyle = xlContinuous
            .Weight = xlMedium
            .Color = RGB(0, 0, 0)
        End With
        
        ' === ������ ������� ===
        .columns("A").ColumnWidth = 25  ' ������������
        .columns("B").ColumnWidth = 12  ' ������
        .columns("C").ColumnWidth = 10  ' ����������
        .columns("D").ColumnWidth = 8   ' �������
        .columns("E").ColumnWidth = 20  ' �����
        
        ' === ������������ ������ ===
        .rows("1:3").AutoFit
        
        ' === ������� ������ ===
        .PageSetup.PrintArea = "A1:E" & (6 + dataCount + 2)
        
        ' === ��������� � ������ ===
        .PageSetup.CenterHeader = "&B&14����� � ���������� �������"
        .PageSetup.RightFooter = "&D &T - ���. &P �� &N"
        
    End With
    
    Exit Sub
    
ErrorHandler:
    MsgBox "������ ��������������: " & Err.description, vbCritical
End Sub

' ===== ��������������� ������� =====

Private Function GetActiveWarehouseCount() As Integer
    Dim count As Integer
    count = 0
    
    If CheckBox1.value Then count = count + 1   ' ����� 5
    If CheckBox2.value Then count = count + 1   ' ����� 6
    If CheckBox3.value Then count = count + 1   ' ����� 7
    If CheckBox4.value Then count = count + 1   ' ����� 8
    If CheckBox5.value Then count = count + 1   ' ����� 9
    If CheckBox6.value Then count = count + 1   ' ����� 10
    If CheckBox7.value Then count = count + 1   ' ����� 11
    If CheckBox8.value Then count = count + 1   ' ����� 12
    
    GetActiveWarehouseCount = count
End Function

Private Function GetActiveWarehousesList() As String
    Dim list As String
    list = ""
    
    If CheckBox1.value Then list = list & "5, "
    If CheckBox2.value Then list = list & "6, "
    If CheckBox3.value Then list = list & "7, "
    If CheckBox4.value Then list = list & "8, "
    If CheckBox5.value Then list = list & "9, "
    If CheckBox6.value Then list = list & "10, "
    If CheckBox7.value Then list = list & "11, "
    If CheckBox8.value Then list = list & "12, "
    
    ' ������� ��������� �������
    If Len(list) > 2 Then
        list = Left(list, Len(list) - 2)
    End If
    
    GetActiveWarehousesList = list
End Function

Private Sub UpdateWarehouseStatus()
    Dim activeCount As Integer
    activeCount = GetActiveWarehouseCount()
    
    If activeCount = 0 Then
        Label3.Caption = "?? �� ������ �� ���� �����! �������� ������ ��� ����������."
        Label3.BackColor = RGB(255, 200, 200)
    ElseIf activeCount = 8 Then
        Label3.Caption = "? ��� ������ ������� - ����� � ����������"
        Label3.BackColor = RGB(200, 255, 200)
    Else
        Label3.Caption = "?? ������� �������: " & activeCount & " �� 8 (" & GetActiveWarehousesList() & ")"
        Label3.BackColor = RGB(200, 230, 255)
    End If
End Sub

Private Sub SyncWarehouseStates()
    Call ModuleWarehouseCapacity.SetWarehouseActive("5", CheckBox1.value)
    Call ModuleWarehouseCapacity.SetWarehouseActive("6", CheckBox2.value)
    Call ModuleWarehouseCapacity.SetWarehouseActive("7", CheckBox3.value)
    Call ModuleWarehouseCapacity.SetWarehouseActive("8", CheckBox4.value)
    Call ModuleWarehouseCapacity.SetWarehouseActive("9", CheckBox5.value)
    Call ModuleWarehouseCapacity.SetWarehouseActive("10", CheckBox6.value)
    Call ModuleWarehouseCapacity.SetWarehouseActive("11", CheckBox7.value)
    Call ModuleWarehouseCapacity.SetWarehouseActive("12", CheckBox8.value)
End Sub

Private Sub SaveWarehouseSettings()
    On Error Resume Next
    SaveSetting "SmartWarehouse", "ActiveWarehouses", "Warehouse5", IIf(CheckBox1.value, "1", "0")
    SaveSetting "SmartWarehouse", "ActiveWarehouses", "Warehouse6", IIf(CheckBox2.value, "1", "0")
    SaveSetting "SmartWarehouse", "ActiveWarehouses", "Warehouse7", IIf(CheckBox3.value, "1", "0")
    SaveSetting "SmartWarehouse", "ActiveWarehouses", "Warehouse8", IIf(CheckBox4.value, "1", "0")
    SaveSetting "SmartWarehouse", "ActiveWarehouses", "Warehouse9", IIf(CheckBox5.value, "1", "0")
    SaveSetting "SmartWarehouse", "ActiveWarehouses", "Warehouse10", IIf(CheckBox6.value, "1", "0")
    SaveSetting "SmartWarehouse", "ActiveWarehouses", "Warehouse11", IIf(CheckBox7.value, "1", "0")
    SaveSetting "SmartWarehouse", "ActiveWarehouses", "Warehouse12", IIf(CheckBox8.value, "1", "0")
End Sub

Private Sub LoadWarehouseSettings()
    On Error Resume Next
    CheckBox1.value = (GetSetting("SmartWarehouse", "ActiveWarehouses", "Warehouse5", "1") = "1")
    CheckBox2.value = (GetSetting("SmartWarehouse", "ActiveWarehouses", "Warehouse6", "1") = "1")
    CheckBox3.value = (GetSetting("SmartWarehouse", "ActiveWarehouses", "Warehouse7", "1") = "1")
    CheckBox4.value = (GetSetting("SmartWarehouse", "ActiveWarehouses", "Warehouse8", "1") = "1")
    CheckBox5.value = (GetSetting("SmartWarehouse", "ActiveWarehouses", "Warehouse9", "1") = "1")
    CheckBox6.value = (GetSetting("SmartWarehouse", "ActiveWarehouses", "Warehouse10", "1") = "1")
    CheckBox7.value = (GetSetting("SmartWarehouse", "ActiveWarehouses", "Warehouse11", "1") = "1")
    CheckBox8.value = (GetSetting("SmartWarehouse", "ActiveWarehouses", "Warehouse12", "1") = "1")
    Call SyncWarehouseStates
End Sub

Private Function CountSuccessful(report As String) As Integer
    Dim lines() As String
    lines = Split(report, vbNewLine)
    
    Dim count As Integer
    count = 0
    
    Dim i As Integer
    For i = 0 To UBound(lines)
        If InStr(lines(i), "?") > 0 Then
            count = count + 1
        End If
    Next i
    
    CountSuccessful = count
End Function

