Attribute VB_Name = "��Ʈ����"
Sub mkFile()
 
    Dim inTemp As Variant
    Dim vrTemp As Variant
    Dim oldBook As Workbook
    Dim newBook As Workbook
    Dim i As Integer
    
    'ó�� �ӵ� ���̱�
    With Application
        .Calculation = xlCalculationManual
        .EnableEvents = False
        .ScreenUpdating = False
    End With
    
    On Error GoTo j '�������� ��ũ�� ����
    
    '���� ��ũ��Ʈ ����
    Set oldBook = ActiveWorkbook
    
    'inputbox�� �̿��Ͽ� ������ ��Ʈ��ȣ ����
    inTemp = InputBox("������ ���� �Է��ϼ���" & vbCr & vbCr & "��) 1,2,3")
    If TypeName(inTemp) = "Boolean" Or inTemp = "" Then GoTo j
    
    '�޸��� �������� ���� �� �и�
    vrTemp = Split(inTemp, ",")
    
    '��Ʈ��ȣ�� �޾� �����Ϸ� ����
    For i = 1 To UBound(vrTemp) + 1
        If i = 1 Then
            Sheets(Val(vrTemp(i - 1)) + 2).Copy
            Set newBook = ActiveWorkbook
            
        Else
            oldBook.Sheets(Val(vrTemp(i - 1)) + 2).Copy after:=newBook.Sheets(newBook.Sheets.Count)
        End If
    Next i
    
j:
    'ó���ӵ� ���� �� ������ ����
    With Application
        .Calculation = xlCalculationAutomatic
        .EnableEvents = True
        .ScreenUpdating = True
    End With
    
End Sub

