Attribute VB_Name = "z_TTS"
'###############################################################
'�����ο��� VBA ����������Լ� (https://www.oppadu.com)
'�� VocSpeak �Լ�
'�� �ܾ Ư�� ���� �о��ִ� �ؽ�Ʈ ���� ��ȯ �Լ��Դϴ�.
'�� �μ� ����
'_____________Voca                  : ���� ��ȯ �� �ؽ�Ʈ�Դϴ�.
'_____________Language          : �ؽ�Ʈ�� ���� ��� ����Դϴ�. ������� PC�� �ش� ������� ��ġ�Ǿ� �־�� �մϴ�.
'###############################################################
 
Public Voc As Object
 
Sub VocSpeak(Optional Voca As Variant, Optional Language As String, Optional blnStop As Boolean = False)
 
' ���� ����
If IsMissing(Voca) Then Voca = ""
Set Voc = CreateObject("SAPI.SpVoice")

' ������ ���� ��ȯ ����� �ϳ��� ���ư��� ��� ���� ��ġ�ϴ� �׸��� �ִ��� Ȯ��
For i = 0 To Voc.GetVoices.Count - 1
    Set Voc.Voice = Voc.GetVoices.Item(i)
    If InStr(1, Voc.Voice.GetDescription, Language) Then GoTo Speak
Next
 
'��ġ�ϴ� �׸��� ���� ��� �ȳ��޽��� ���� �Լ� ����
MsgBox "���� ��ȯ �� �� PC�� ��ġ�Ǿ� ���� �ʽ��ϴ�.", vbInformation, "�����ο��� - �����ȳ�"
 
Exit Sub
 

' ������ ���� ��ȯ �׸��� ���� �� ���� ��ȯ �� ��ɹ��� �����մϴ�.
Speak:
Voc.Speak Voca, 1
If blnStop = True Then Voc.Speak "", 0

End Sub

'###############################################################
'�����ο��� VBA ����������Լ� (https://www.oppadu.com)
'�� GetVocList �Լ�
'�� ��� ������ ���� ��ȯ ����� �迭�� ���l���ϴ�.
'�� �μ� ����
'_____________ShowMsgBox : True�� ��� ��밡���� ���� ��ȯ ����� �޽����ڽ��� ����մϴ�.
'###############################################################
Function GetVocList(Optional ShowMsg As Boolean = False) As Variant
 
' ���� ����
Dim Voc As Object
Dim vaReturn As Variant
Dim v As Variant: Dim s As String
Set Voc = CreateObject("SAPI.SpVoice")
 
' ��� ������ ���� ��ȯ ��� ���� ���̷� �迭 ����
ReDim vaReturn(0 To Voc.GetVoices.Count - 1)
 
' ���� �̸��� �迭�� �߰�
For i = 0 To Voc.GetVoices.Count - 1
    Set Voc.Voice = Voc.GetVoices.Item(i)
    vaReturn(i) = Voc.Voice.GetDescription
Next
 
' �޼��� ��¿��� True�� ��� �޼��� ���
If ShowMsg = True Then: For Each v In vaReturn: s = s & v & vbNewLine: Next: MsgBox s
 
' ����� ��ȯ
GetVocList = vaReturn
 
End Function
