Attribute VB_Name = "z_PageSetup"
Public Enum ePrintMargin
    xlNone = 0
    xlNarrow = 1
    xlNormal = 2
    xlWide = 3
End Enum

Public Enum ePaperSize
    xlA4 = 9
    xlA3 = 8
    xlLetter = 1
    xlA5 = 11
End Enum

Function getPrintMargin(eValue As ePrintMargin) As Variant

'// ������ eNum ������ ������ ���鼳���� ���� ���� �迭�� �����մϴ�.

Select Case eValue
    Case 0
        getPrintMargin = Array(0.05, 0.05, 0.05, 0.05, 0.1, 0.1)
    Case 1
        getPrintMargin = Array(0.25, 0.25, 0.75, 0.75, 0.3, 0.3)
    Case 2
        getPrintMargin = Array(0.7, 0.7, 0.75, 0.75, 0.3, 0.3)
    Case 3
        getPrintMargin = Array(1, 1, 1, 1, 0.5, 0.5)
End Select

End Function

Sub Page_Setup(WS As Worksheet, Optional LHead As String = "", Optional RHead As String = "&D / &T", _
                Optional LFoot As String = "�� �������� ���ܺ����� ���մϴ�.", Optional RFoot As String = "&P / &N ������", _
                Optional eMargin As ePrintMargin = xlNarrow, _
                Optional HFit As Boolean = True, Optional VFit As Boolean = False, _
                Optional HCenter As Boolean = True, Optional VCenter As Boolean = False, _
                Optional eOrient As XlPageOrientation = xlPortrait, Optional eSize As ePaperSize = xlA4)

Dim pSetup As String
Dim varMargin As Variant
Dim lngOrient As Integer

'// �μ⼳�� ������Ʈ �ߴ� (�ӵ�����)
Application.PrintCommunication = False

'// �μ⿩�鰪�� �޾ƿɴϴ�.
varMargin = getPrintMargin(eMargin)

'// �μ���� ������ �����մϴ�.
If eOrient = xlPortrait Then
    lngOrient = 1
Else
    lngOrient = 2
End If

'// ExecuteExcel4Macro �� Page.Setup ��ɹ� ������ ���� ������ �Է��մϴ�.
Head = """&L" & LHead & "&R" & RHead & """"     '// ������ �Ӹ����Դϴ�.
Foot = """&L" & LFoot & "&R" & RFoot & """"     '// ������ �������Դϴ�.
pLeft = varMargin(0)                            '// ���ʿ���
pRight = varMargin(1)                           '// �����ʿ���
Top = varMargin(2)                              '// ������
Bot = varMargin(3)                              '// �Ʒ�����
Head_margin = varMargin(4)                      '// �Ӹ�������
Foot_margin = varMargin(5)                      '// ����������
Hdng = 0                                        '// ��/���ݺ� ��¿��� 0 = �ݺ���¾��� 1 = �ݺ����
Grid = False                                    '// ���ݼ���¿���
Notes = False                                   '// �޸���¿���
H_cntr = HCenter                                '// �������
V_cntr = VCenter                                '// �߾�����
Orient = lngOrient                              '// ��������, 1 = ���� 2 = ����
Paper_size = eSize                              '// ����ũ��
Pg_num = 1                                      '// ������ ���۹�ȣ
Pg_order = 1                                    '// ��������ȣ ����, 1 = ��-�Ʒ�-�� 2 = ��-��-�Ʒ�
Quality = ""                                    '// �μ�ǰ�� (dot-per-inch�� �Է�) (���� = �ڵ�)
bw_cells = False                                '// ����μ⿩��, TRUE = ����/�׵θ� ����,��� ��� FALSE = ����
pScale = 100                                    '// ���/Ȯ����� �Ǵ� TRUE (Fit to Page)

'// ������ �������� ������ ��� �Ӹ���/�������� �����Ͽ� �μ⿵���� ��ġ�� �ʵ��� �մϴ�.
If eMargin = xlNone Then
    Head = """"""
    Foot = """"""
End If


'// ExecuteExcel4Macro ��ɹ��� �����մϴ�.
pSetup = "PAGE.SETUP(" & Head & ", " & Foot & ", " & pLeft & ", " & pRight & ", " & Top & ", " & Bot & ", "
pSetup = pSetup & Hdng & ", " & Grid & "," & H_cntr & "," & V_cntr & "," & Orient & ","
pSetup = pSetup & Paper_size & "," & pScale & ","
pSetup = pSetup & Pg_num & "," & Pg_order & "," & bw_cells & "," & Quality & ","
pSetup = pSetup & Head_margin & "," & Foot_margin & "," & Notes & ")"


Application.ExecuteExcel4Macro pSetup

'// ExecuteExcel4Macro������ '�� �������� ��/�� ���߱�' ����� �������� �ʽ��ϴ�.
'// ���� ��Ʈ�� PageSetup �Ӽ����� '������ ��/�� ���߱� ����� �����մϴ�.
With WS.PageSetup
    If HFit = True Then
        .FitToPagesWide = 1
    Else
        .FitToPagesWide = False
    End If
    
    If VFit = True Then
        .FitToPagesTall = 1
    Else
        .FitToPagesTall = False
    End If
End With

'// �μ⼳�� ������Ʈ
Application.PrintCommunication = True

End Sub

