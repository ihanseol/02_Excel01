Attribute VB_Name = "Module1"
Option Explicit

Sub Web_scarping()

Dim MyBrowser As InternetExplorer
Dim HTMLDoc As HTMLDocument
Dim iArticle As IHTMLElement
Dim i As Long

'// �ͽ��÷ξ�(XML���� ����)
Set MyBrowser = Sheet1.WebBrowser1

'// �˻� (XML ��û)
With MyBrowser
    .Navigate Sheet1.Range("C3").Value
    
    Wait_Browser MyBrowser

    '// HTML ����
    Set HTMLDoc = .Document
    
    i = 7
    
    '// ��ü�˻�
    For Each iArticle In HTMLDoc.getElementsByClassName("_sp_each_title")
        '// ��Һ� �� ����
        Sheet1.Cells(i, 2) = iArticle.Title
        Sheet1.Cells(i, 3) = iArticle.getAttribute("href")
        Sheet1.Hyperlinks.Add Sheet1.Cells(i, 3), Sheet1.Cells(i, 3)
        i = i + 1
    Next
    
End With

MsgBox "�˻��Ͻ� �ܾ� [" & Sheet1.Range("C2") & "] �� ���� ���̹� ������� ��ũ������ �Ϸ��Ͽ����ϴ�."


End Sub

Sub Wait_Browser(Browser As InternetExplorer, Optional t As Integer = 1)

While Browser.Busy Or Browser.ReadyState <> 4
DoEvents
Wend

Application.Wait DateAdd("s", t, Now)

End Sub
