Attribute VB_Name = "Email_Module"
Option Explicit

'##########################################################
' �����ο��� �� VBA 7��, ���� ������ �ڵ�ȭ �ϼ�����
' ��ɹ��� ���� �ڼ��� ������ �Ʒ� ��ũ���� Ȯ���ϼ���.
' https://www.oppadu.com/����-����-������-�ƿ���-��ũ��/
'##########################################################

Sub Test()

Dim FileName As String
Dim SavePath As String

FileName = Sheet4.Range("E3").Value & "_" & Sheet4.Range("C4").Value & "_" & Sheet4.Range("H3").Value & "�� " & Sheet4.Range("H4").Value & "��"
SavePath = GetDesktopPath

Rng_To_Pdf Sheet4.Range("B2:E18"), FileName, SavePath, OpenPdf:=False, AddSequence:=False
'// Rng_To_Pdf ��ɹ��� ù��° �μ��� Selection���� �����ϸ� ������ ������ PDF���Ϸ� �����Ͽ� ���Ͽ� ÷���մϴ�.

Sheet4.Range("B2:E18").Select
'// ���õ� ������ �ƴ� ���ϴ� �κ��� �����ؼ� ���Ͽ� ÷���ϴ� ����� �ñ��Ͻź��� Send_Email ��ɹ� ���� ����Ʈ�� �����ϼ���.

Send_Email "test@oppadu.com", _
            FileName, _
            "<p><span style=""font-family: NanumGothic, �������, sans-serif; font-size: 9pt;""><b><u><span style=""font-size: 9pt;"">������ �븮</span></u></b> �Բ�&nbsp;</span></p><p><br></p><p><span style=""font-family: NanumGothic, �������, sans-serif; font-size: 9pt;"">������ <b><u><span style=""font-size: 9pt;"">2019�� 10�� �޿�����</span></u></b>�� �ۺε帳�ϴ�.&nbsp;</span></p><p><span style=""font-family: NanumGothic, �������, sans-serif; font-size: 9pt;"">�����ο����� ���� ������ ��� ���� ����帮�� ���� ������ ������� ���Ͽ� ��� �����ϰڽ��ϴ�.</span></p>", _
            True, _
            "", , _
            SavePath & FileName & ".pdf" & "|" & SavePath & FileName & ".pdf"

End Sub

'######################################################################
' ��ɹ�    : Send_Email
' ����      : �ƿ���� �����Ͽ� ���Ϻ����⸦ �ڵ�ȭ�ϴ� ����Դϴ�.
' ��ɹ��� ���� �ڼ��� ������ �����ο��� Ȩ�������� �����ϼ���.
' https://www.oppadu.com
'######################################################################

Sub Send_Email(MailTo As String, _
                Subject As String, _
                HTMLString As String, _
                Optional PasteSelection As Boolean = False, _
                Optional CCTo As String = "", _
                Optional BCCTo As String = "", _
                Optional AttachFilePath As String = "", _
                Optional PathDelimiter As String = "|")

Dim AppOutlook As Outlook.Application       '// �ƿ��� ���α׷�
Dim newEmail As Outlook.MailItem            '// �ƿ��� ���� ������ ������ ���� ������ ����
Dim pageInspector As Outlook.Inspector      '// �ƿ��� ���忡���� ������������ �׸�
Dim pageEditor As Object                    '// �ƿ��� �̸��� ����â
Dim varFilePath As Variant                  '// ���ϰ�θ� �迭���·� ������� ����
Dim FileCount As Long                       '// ÷�������� ����
Dim i As Long                               '// For�� �ݺ����� ����
Dim wdPasteDefault As Variant

Set AppOutlook = New Outlook.Application
Set newEmail = AppOutlook.CreateItem(olMailItem)


If AttachFilePath <> "" Then
        varFilePath = Split(AttachFilePath, PathDelimiter)
End If


With newEmail
    .To = MailTo
    .CC = CCTo
    .BCC = BCCTo
    .Subject = Subject
    If AttachFilePath <> "" Then
        For i = 1 To UBound(varFilePath) + 1
            .Attachments.Add varFilePath(i - 1), 1, i
        Next
    End If
    .HTMLBody = HTMLString
    
    '.DeferredDeliveryTime = DateAdd("n", 5, Now)
    .DeferredDeliveryTime = DateSerial(2030, 1, 1) + TimeSerial(8, 0, 0)
    
    If PasteSelection = True Then
        .Display
        Set pageInspector = newEmail.GetInspector
        Set pageEditor = pageInspector.WordEditor
        pageEditor.Application.Selection.Start = Len(.Body)
        
        Selection.Copy
        pageEditor.Application.Selection.PasteAndFormat wdPasteDefault
    Else
        .Display
    End If
    
'.Send    '// ������ �������� �ּ�ó���� �����ϼ���.

End With

Set pageEditor = Nothing
Set pageInspector = Nothing
Set newEmail = Nothing
Set AppOutlook = Nothing

End Sub
