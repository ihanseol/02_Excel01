Attribute VB_Name = "z_Scrape"
Function GetHttp(URL As String, Optional formText As String, Optional isWinHttp As Boolean = False, Optional RequestHeader As Variant) As Object
 
'###############################################################
'�����ο��� VBA ����������Լ� (https://www.oppadu.com)
'�� GetHttp �Լ�
'�� ������ �����͸� �޾ƿɴϴ�.
'�� �μ� ����
'_____________URL                      : �����͸� ��ũ���� �� ������ �ּ��Դϴ�.
'_____________formText              : Encoding �� FormText �������� ������ �� ���, Send String�� �������� �߰��մϴ�.
'_____________isWinHttp             : WinHTTP �� ��û���� �����Դϴ�. Redirect�� �ʿ��� ��� True�� �Է��Ͽ� WinHttp ��û�� �����մϴ�.
'_____________RequestHeader     : RequestHeader�� �迭�� �Է��մϴ�. �ݵ�� ¦��(�� �־� �̷����) ������ �ԷµǾ�� �մϴ�.
'�� ��� ����
'Dim HtmlResult As Object
'Set htmlResult = GetHttp("https://www.naver.com")
'msgbox htmlResult.body.innerHTML
'###############################################################
 
Dim oHTMLDoc As Object: Dim objHTTP As Object
Dim i As Long: Dim blnAgent As Boolean: blnAgent = False
Dim sUserAgent As String: sUserAgent = "Mozilla/5.0 (Linux; Android 6.0; Nexus 5 Build/MRA58N) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/86.0.4240.183 Mobile Safari/537.36"
 
Application.DisplayAlerts = False
 
If Left(URL, 4) <> "http" Then URL = "http://" & URL
 
Set oHTMLDoc = CreateObject("HtmlFile")
 
If isWinHttp = False Then
    Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP")
Else
    Set objHTTP = CreateObject("WinHttp.WinHttpRequest.5.1")
End If
 
objHTTP.setTimeouts 3000, 3000, 3000, 3000
objHTTP.Open "POST", URL, False
If Not IsMissing(RequestHeader) Then
    Dim vRequestHeader As Variant
    For Each vRequestHeader In RequestHeader
        Dim uHeader As Long: Dim Lheader As Long: Dim steps As Long
        uHeader = UBound(vRequestHeader): Lheader = LBound(vRequestHeader)
        If (uHeader - Lheader) Mod 2 = 0 Then GetHttp = CVErr(xlValue): Exit Function
        For i = Lheader To uHeader Step 2
            If vRequestHeader(i) = "User-Agent" Then blnAgent = True
            objHTTP.setRequestHeader vRequestHeader(i), vRequestHeader(i + 1)
        Next
    Next
End If
If blnAgent = False Then objHTTP.setRequestHeader "User-Agent", sUserAgent

objHTTP.send formText
 
With oHTMLDoc
    .Open
    .Write objHTTP.responsetext
    .Close
End With
 
Set GetHttp = oHTMLDoc
Set oHTMLDoc = Nothing
Set objHTTP = Nothing
 
Application.DisplayAlerts = True
 
End Function

Function Splitter(v As Variant, Cutter As String, Optional Trimmer As String)
 
'###############################################################
'�����ο��� VBA ����������Լ� (https://www.oppadu.com)
'�� Splitter �Լ�
'�� Cutter ~ Timmer ������ ���ڸ� �����մϴ�. (Timmer�� ��ĭ�� ��� Cutter ���� ���ڿ��� �����մϴ�.)
'�� �μ� ����
'_____________v       : ���ڿ��Դϴ�.
'_________Cutter      : ���ڿ� ������ ������ �ؽ�Ʈ�Դϴ�.
'_________Trimmer  : ���ڿ� ������ ������ �ؽ�Ʈ�Դϴ�. (�����μ�)
'�� ��� ����
'Dim s As String
's = "{sa;b132@drama#weekend;aabbcc"
's = Splitter(s, "@", "#")
'msgbox s   '--> "drama"�� ��ȯ�մϴ�.
'###############################################################
 
Dim vaArr As Variant

On Error GoTo EH:

vaArr = Split(v, Cutter)(1)
If Not IsMissing(Trimmer) Then vaArr = Split(vaArr, Trimmer)(0)
 
Splitter = vaArr

Exit Function

EH:
    Splitter = ""
    
End Function
