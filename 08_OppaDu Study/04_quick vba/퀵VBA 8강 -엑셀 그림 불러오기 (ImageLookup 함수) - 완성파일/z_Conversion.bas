Attribute VB_Name = "z_Conversion"
Option Explicit

'###############################################################
'�����ο��� VBA ����������Լ� (https://www.oppadu.com)
'�� cvRng �Լ�
'�� ����������Լ����� ������ �μ��� ���� �� ����մϴ�. ���� �μ��� ������ �ԷµǾ��� ���, ������ �Էµ� ���� ��ȯ�մϴ�.
'�� �μ� ����
'_____________TargetRng     : ���� ��ȯ�� ���� �Ǵ� �׿� ���Դϴ�.
'###############################################################
Function cvRng(TargetRng)

If TypeName(TargetRng) = "Range" Then
    cvRng = TargetRng.Value
Else
    cvRng = TargetRng
End If

End Function
