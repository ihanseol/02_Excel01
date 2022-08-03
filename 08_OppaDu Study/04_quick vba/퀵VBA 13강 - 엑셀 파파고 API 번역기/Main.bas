Attribute VB_Name = "Main"
Option Explicit

Sub DoTranslate()

'-----------------------------------------
' 1. ���� ���� (Rng, r)
'-----------------------------------------
Dim Rng As Range            '���� �� �Էµ� ����
Dim r As Range                  'Rng �� �ϳ��� ���ư��鼭 ������ ����

'-----------------------------------------
' 2. ���� �����Ǳ� ��/�� ���� ���� ��� �ʱ�ȭ
'-----------------------------------------
On Error Resume Next
Sheet2.Range("tblResult").Delete
On Error GoTo 0

'-----------------------------------------
' 3. ���� ������ �Էµ� ���� ����
'-----------------------------------------
Set Rng = Sheet2.Range("tblOriginal")

'-----------------------------------------
' 4. ������ �ϳ��� ���ư��� Papago ���� ���� -> ������� ��ȯ
'-----------------------------------------
For Each r In Rng
    Sheet2.Range("P" & r.Row).Value = PapagoTranslate(r.Value)
Next

'-----------------------------------------
' 5. ���İ� API ��뷮 ������Ʈ
'-----------------------------------------
Sheet2.Range("U6").Value = Sheet2.Range("U6").Value + Sheet2.Range("F11").Value

MsgBox "���� ������ �Ϸ�Ǿ����ϴ�!"

End Sub

Sub asdlf()



End Sub
'=VLOOKUP(ã����,����,����ȣ...)
'=PapagoTranslate(����)
Function PapagoTranslate(OriginalText) As String

'-----------------------------------------
'1. ���� �����ϱ�
'-----------------------------------------
Dim sID As String           ' API Ű, ����Ű
Dim sSecret As String
Dim sFrom As String         ' ���� ���
Dim sTo As String           ' ���� ���

Dim URL As String           ' API URL ("https://openapi.naver.com/v1/papago/n2mt")
Dim Query As String         ' ���� ��Ʈ�� ("source=ko&target=en&text=�ȳ��ϼ���. ������ �ݰ����ϴ�.")
Dim vArray As Variant       ' HTTP ��û�� ���� Request Header
Dim objHTML As HTMLDocument ' HTTP ��û���� ��ȯ�� HTML ����
Dim sResult As String       ' HTTML ��û���� ��ȯ�� HTML �ؽ�Ʈ

'-----------------------------------------
'2. ���� ����
'-----------------------------------------
'APIŰ�� ����Ű�� ����
sID = Sheet3.Range("B1").Value
sSecret = Sheet3.Range("B2").Value

'�������� ��������� �ڵ� ã�� (VLOOKUP �Լ�)
sFrom = Application.WorksheetFunction.VLookup(Sheet2.Range("F9").Value, Sheet3.Range("A:B"), 2, 0)
sTo = Application.WorksheetFunction.VLookup(Sheet2.Range("P9").Value, Sheet3.Range("A:B"), 2, 0)

''HTTP ��û�� ���� URL�� Request Header �����ϱ�
URL = "https://openapi.naver.com/v1/papago/n2mt"
Query = "source=" & sFrom & "&target=" & sTo & "&text=" & OriginalText

'Request Header�� �� 4��, Uger-Agent, Content-Type, Client-Id, Client-Secret
ReDim vArray(0 To 3)
vArray(0) = Array("User-Agent", "curl/749.1")
vArray(1) = Array("Content-Type", "application/x-www-form-urlencoded; charset=UTF-8")
vArray(2) = Array("X-Naver-Client-Id", CStr(sID))
vArray(3) = Array("X-Naver-Client-Secret", CStr(sSecret))

'' ������ �Ʒ��� ���� �Էµǰ���? :)
''"source=ko&target=en&text=�ȳ��ϼ���. ������ �ݰ����ϴ�."
''���� ���� ������ ������ ��ɹ��� �ۼ��մϴ�.
'-----------------------------------------
'3. HTTP ��û���� API ��� �޾ƿ���
'-----------------------------------------
Set objHTML = GetHttp(URL, Query, False, vArray)
sResult = objHTML.body.innerHTML

'-----------------------------------------
'4. Splitter�� ���ϴ� ����� �����ϱ�
'-----------------------------------------
sResult = Splitter(sResult, "translatedText"":""", """,""engineType")

'������ �߻��� ��� �����޽����� ��ȯ
If sResult = "" Then sResult = Splitter(sResult, "errorMessage"":""", """,""errorCode")

'-----------------------------------------
'5. PapagoTranslate �Լ��� ��� ��ȯ �� ��ɹ� ����
'-----------------------------------------
PapagoTranslate = sResult

End Function
