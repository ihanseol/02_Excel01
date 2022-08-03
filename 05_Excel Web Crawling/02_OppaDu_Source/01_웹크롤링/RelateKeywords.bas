Attribute VB_Name = "RelateKeywords"
Sub deleteSearchKeywords()

    Sheet2.Range("f2") = ""

End Sub


Sub SearchRelateKeywords()

Dim url As String
Dim htmlResult As Object
Dim strResult As String
Dim strSearch As String

Dim v As Variant
Dim i As Long



'strSearch = "ȭ��"
strSearch = Sheet2.Range("f2")
If Not HasContent(strSearch) Then strSearch = "���"

url = "https://ac.search.naver.com/nx/ac?q=" & strSearch & "&frm=nv&st=100"

Set htmlResult = GetHttp(url)

strResult = htmlResult.body.innerHTML

'ExportText strResult

'splitter - in json format
'{ "query" : ["ȭ��"], "items" : [ [["ȭ����"],["ȭ����û"],["ȭ��"],["ȭ�� �ڷγ�"],["ȭ�� ����"],["ȭ������"],["ȭ�������뺴��"],["ȭ�� ��纴��"],["ȭ�� �ڿ��ǹ���"],["ȭ�� ��纴�� �ڷγ�"]] ] }
'start : <"items" : [ [> , end : <] ] }>
'�߶󳽴� ���۰� ���� ���̽� ���˽�Ʈ����

Debug.Print strResult
strResult = Splitter(strResult, """items"" : [ [", "] ] }")

v = Split(strResult, ",")


On Error GoTo EmptyStrResult
For i = LBound(v) To UBound(v)
    v(i) = Replace(v(i), "[""", "")
    v(i) = Replace(v(i), """]", "")
    Debug.Print (v(i))
Next i


ArrayToRng Sheet2.Range("A1"), v

EmptyStrResult:

End Sub



