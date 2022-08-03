Attribute VB_Name = "�޷�"
Option Explicit
Sub Calendar()
'������ �޷� �����

    Dim i As Integer
    Dim Days As Integer, StartDay As Integer
    Dim Row As Integer, cnt As Integer
    Dim InputYear As Integer
    Dim theDay As Date
    Dim formSheet As Worksheet

    With Application
        .ScreenUpdating = False
        .EnableEvents = False
        .Calculation = xlCalculationManual
    
        '���� ��ũ��Ʈ ����
        .DisplayAlerts = False
            For i = 3 To Sheets.Count
                Sheets(3).Delete
            Next i
        .DisplayAlerts = True
    End With
    
    '�������� ��ũ�� ����
    On Error GoTo j
    
    Set formSheet = Sheets("����")  '���Ľ�Ʈ
    formSheet.Visible = True   '���� ��Ʈ ���̰�
    InputYear = CInt(InputBox("�� �⵵ �޷��� ���弼��??", "�޷� �����", year(Date)))
    
    For i = 1 To 12
        
        '���Ľ�Ʈ���� ��Ʈ���� �� ��Ʈ�� ����
        formSheet.Copy after:=Sheets(Sheets.Count)
        Sheets(Sheets.Count).Name = i & "��"
        
        '���� �޷� �����
        Range("B1") = InputYear '���� ���
        Range("C1") = i '��
        Days = DateSerial(InputYear, i + 1, 1) - DateSerial(InputYear, i, 1)    '�� ���� �� ��¥
        StartDay = Weekday(DateSerial(InputYear, i, 1), vbSunday) + 1   '�޷� �����ϴ� ��¥
        Row = 4
        cnt = 1
        
        '�޷� �Է�
        Do
            theDay = DateSerial(InputYear, i, cnt)
            With Cells(Row, StartDay)
                .Value = theDay
        
                '�ָ�, �������� ����������
                With .Font
        
                    '�ָ�
                    Select Case Weekday(theDay)
                        Case 1, 7: .Color = vbRed
                    End Select
            
                    '��°�����
                    Select Case Format(theDay, "m.d")
                        Case "1.1", "3.1", "5.5", "6.6", "8.15", "10.3", "10.9", "12.25"
                        .Color = vbRed
                    End Select
                    
                    '���°�����
                    Select Case Sol2Lun(InputYear, i, cnt)
                        Case "1.1", "1.2", "4.8", "8.14", "8.15", "8.16"
                        .Color = vbRed
                    End Select
                    If Sol2Lun(year(theDay + 1), Month(theDay + 1), Day(theDay + 1)) = "1.1" Then: .Color = vbRed '���� ������ ���⵵ 12.31�̹Ƿ� ���� ����
                    
                    '��ü������
                    Select Case theDay
                        Case #9/29/2015#, #2/10/2016#, #1/30/2017#, #9/26/2018#, #5/7/2018#, #5/6/2019#, #1/27/2020#, #9/12/2022#, #1/24/2023#, #2/12/2024#, #5/6/2024#, #10/8/2025#, #2/9/2027#, #9/24/2029#, #5/7/2029#
                            .Color = vbRed
                    End Select
                End With
            End With
        
            StartDay = StartDay + 1
            cnt = cnt + 1
            If StartDay = 9 Then
                StartDay = 2
                Row = Row + 2
            End If
        Loop While cnt <= Days
        
    Next i

    Sheets(3).Activate
    
j:
    formSheet.Visible = False
    With Application
        .ScreenUpdating = True
        .EnableEvents = True
        .Calculation = xlCalculationAutomatic
    End With
End Sub

Sub ExportSchedule()

    RefChkBox.Show

End Sub

Sub ClearCalendar()
'�޷� �ʱ�ȭ

    Dim i As Integer
    
    Application.ScreenUpdating = False
    
    If Sheets.Count > 2 Then
        
        For i = 3 To Sheets.Count
            With Sheets(i)
                .Rows(5).ClearContents
                .Rows(7).ClearContents
                .Rows(9).ClearContents
                .Rows(11).ClearContents
                .Rows(13).ClearContents
                .Rows(15).ClearContents
            End With
            
        Next
    
    End If
    
    Application.ScreenUpdating = True
    
End Sub
