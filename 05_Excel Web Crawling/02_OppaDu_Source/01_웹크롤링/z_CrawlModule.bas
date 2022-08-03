Attribute VB_Name = "z_CrawlModule"

Public Function HasContent(strName As String) As Boolean
    HasContent = (Len(Trim(strName)) > 0)
End Function



Function GetLastRow(WS As Worksheet) As Long

'###############################################################
'�����ο��� VBA ����������Լ� (https://www.oppadu.com)
'�� GetLastRow �Լ�
'�� ��Ʈ���� ���� ������ ���� ��ȯ�մϴ�.
'�� �μ� ����
'_____________WS        : ������ ���� ��ȸ�� ��� ��Ʈ�Դϴ�.
'###############################################################

With WS
GetLastRow = .UsedRange.Rows.Count + .UsedRange.Row - 1
End With

End Function

Sub InsertWebImage(targetRng As Range, ImgLink As String, Optional imgWidth As Double, Optional imgHeight As Double)

'###############################################################
'�����ο��� VBA ����������Լ� (https://www.oppadu.com)
'�� InsertWebImage �Լ�
'�� �� �̹����� �� ���� �����մϴ�.
'�� �μ� ����
'_____________targetRng                              : �̹����� ������ ���Դϴ�.
'_____________imgLink                                 : ������ �̹��� ��ũ�Դϴ�.
'_____________imgWidth                             : �̹��� �ʺ��Դϴ�. �⺻���� ���� �ʺ��Դϴ�. (�����μ�)
'_____________imgHeight                            : �̹��� �����Դϴ�. �⺻���� ���� �����Դϴ�. (�����μ�)
'�� ��� ����
'InsertWebImage(sheet1.Range("A1"),"https://www.google.com/images/branding/googlelogo/1x/googlelogo_color_272x92dp.png")
'###############################################################

Dim WS As Worksheet
Dim shp As Shape
Dim W As Double: Dim H As Double

Set WS = targetRng.Parent
If IsMissing(imgWidth) Then W = imgWidth Else W = targetRng.Width - 6
If IsMissing(imgHeight) Then W = imgHeight Else H = targetRng.Height - 6

Set shp = WS.Shapes.AddPicture(ImgLink, msoFalse, msoTrue, targetRng.Left + 3, targetRng.Top + 3, W, H)

End Sub

Function GetHttp(url As String, ParamArray RequestHeader() As Variant) As Object

'###############################################################
'�����ο��� VBA ����������Լ� (https://www.oppadu.com)
'�� GetHttp �Լ�
'�� ������ �����͸� �޾ƿɴϴ�.
'�� �μ� ����
'_____________URL                               : �����͸� ��ũ���� �� ������ �ּ��Դϴ�.
'_____________RequestHeader            : RequestHeader�� �迭�� �Է��մϴ�. �ݵ�� ¦��(�� �־� �̷����) ������ �ԷµǾ�� �մϴ�.
'�� ��� ����
'Dim HtmlResult As Object
'Set htmlResult = GetHttp("https://www.naver.com")
'msgbox htmlResult.body.innerHTML
'###############################################################

Dim oHTMLDoc As Object: Dim objHTTP As Object

Dim sUserAgent As String: sUserAgent = "Mozilla/5.0 (Linux; Android 6.0; Nexus 5 Build/MRA58N) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/86.0.4240.183 Mobile Safari/537.36"

Application.DisplayAlerts = False

Set oHTMLDoc = CreateObject("HtmlFile")
Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP")

objHTTP.setTimeouts 3000, 3000, 3000, 3000
objHTTP.Open "GET", url, False
objHTTP.setRequestHeader "User-Agent", sUserAgent
If Not IsMissing(RequestHeader) Then
    Dim vRequestHeader As Variant
    For Each vRequestHeader In RequestHeader
        Dim uHeader As Long: Dim Lheader As Long: Dim steps As Long
        uHeader = UBound(vRequestHeader): Lheader = LBound(vRequestHeader)
        If (uHeader - Lheader) Mod 2 = 0 Then GET_HttpRequest = CVErr(xlValue): Exit Function
        For i = Lheader To uHeader Step 2
            objHTTP.setRequestHeader vRequestHeader(i), vRequestHeader(i + 1)
        Next
    Next
End If
objHTTP.send
        
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

Sub ExportText(InnerStrings As String, Optional fileName As String = "�ؽ�Ʈ����", Optional Path As String)

'###############################################################
'�����ο��� VBA ����������Լ� (https://www.oppadu.com)
'�� Export_Text �Լ�
'�� ���ڿ��� �ؽ�Ʈ���Ϸ� �����մϴ�.
'�� �μ� ����
'_____________InnerStrings      : �ؽ�Ʈ���Ϸ� ������ ���ڿ��Դϴ�.
'_____________fileName           : �ؽ�Ʈ ���� �̸��Դϴ�. �⺻���� "�ؽ�Ʈ����" �Դϴ�. (�����μ�)
'_____________path                   : �ؽ�Ʈ ������ ������ ����Դϴ�. �⺻���� ����ȭ���Դϴ�. (�����μ�)
'�� ��� ����
'ExportText "������ �ؽ�Ʈ"
'###############################################################

Dim TextFile As Integer
Dim FilePath As String

If Not IsMissing(Path) Then Path = Environ("USERPROFILE") & "\Desktop\"
FilePath = Path & fileName & ".txt"

TextFile = FreeFile

Open FilePath For Output As TextFile
Print #TextFile, InnerStrings
Close TextFile

End Sub

Sub ArrayToRng(startRng As Range, Arr As Variant)

'###############################################################
'�����ο��� VBA ����������Լ� (https://www.oppadu.com)
'�� ArrayToRng �Լ�
'�� �迭�� ���� ���� ��ȯ�մϴ�.
'�� �μ� ����
'_____________startRng      : �迭�� ��ȯ�� ���� ����(��) �Դϴ�.
'_____________Arr               : ��ȯ�� �迭�Դϴ�.
'�� ��� ����
'Dim v As Variant
'ReDim v(0 to 1)
''v(0) = "a" : v(1) = "b"
'ArrayToRng Sheet1.Range("A1"), v
'##############################################################

On Error GoTo SingleDimension:
startRng.Cells(1, 1).Resize(UBound(Arr, 1) - LBound(Arr, 1) + 1, UBound(Arr, 2) - LBound(Arr, 2) + 1) = Arr

Exit Sub
SingleDimension:
Dim tempArr As Variant: Dim i As Long
ReDim tempArr(LBound(Arr, 1) To UBound(Arr, 1), 1 To 1)
For i = LBound(Arr, 1) To UBound(Arr, 1)
    tempArr(i, 1) = Arr(i)
Next
startRng.Cells(1, 1).Resize(UBound(Arr, 1) - LBound(Arr, 1) + 1, 1) = tempArr

End Sub

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

vaArr = Split(v, Cutter)(1)
If Not IsMissing(Trimmer) Then vaArr = Split(vaArr, Trimmer)(0)

Splitter = vaArr

End Function

Function ParseJSON(strJSON, strToParse, Optional strID, Optional strToRemove) As Variant

'###############################################################
'�����ο��� VBA ����������Լ� (https://www.oppadu.com)
'�� ParseJSON �Լ�
'�� strJSON �����Ϳ��� ������ ������ ���� �����մϴ�.
'�� �μ� ����
'_____________strJSON            : JSON �������Դϴ�.
'_____________strToParse        : JSON ������ ������ �ʵ���Դϴ�. ��ǥ(,)�� �����Ͽ� �Է��մϴ�.
'_____________strID                 : ������ ������ �迭�� ���� �߰��� ID �Դϴ�. (�����μ�)
'_____________strToRemove   : ������ �����Ϳ��� ������ ���ڿ��Դϴ�. ��ǥ(,)�� �����Ͽ� �Է��մϴ�.
'�� ��� ����
'Dim v As Variant
'v = ParseJSON(JsonData, "Date, Name, Item")
'###############################################################

'----------------------------------------------------
'���� ����
'----------------------------------------------------
Dim vaToParse As Variant: Dim vToParse As Variant
Dim vaToRemove As Variant: Dim vToRemove As Variant
Dim lngStart As Long
Dim objItm As Variant: Dim strItm As String: Dim tmpItm As String
Dim itmCnt As Long
Dim i As Long: Dim r As Long: Dim c As Long
Dim dicItm As Object
Dim vaItm As Variant: Dim vaItems As Variant: Dim vaReturn As Variant
Dim iCol As Long: Dim maxCol As Long: Dim j As Long

Set dicItm = CreateObject("Scripting.Dictionary")

'----------------------------------------------------
'JSON ���� ����
'----------------------------------------------------
vaToParse = Split(strToParse, ",")
If Not IsMissing(strToRemove) Then vaToRemove = Split(strToRemove, ",")

lngStart = InStr(1, strJSON, "[")
strJSON = Right(strJSON, Len(strJSON) - lngStart)

objItm = Split(strJSON, "{""")
itmCnt = UBound(objItm)

For i = 1 To itmCnt
    strItm = Split(objItm(i), """}")(0)
     iCol = Len(strItm) - Len(Replace(strItm, ":", ""))
    If iCol > maxCol Then maxCol = iCol
    
    If Not IsMissing(strID) Then
        ReDim vaItm(0 To iCol)
        vaItm(0) = strID: j = 1
    Else
        ReDim vaItm(0 To iCol - 1)
        j = 0
    End If
    
    On Error Resume Next
    For Each vToParse In vaToParse
        If InStr(strItm, Trim(vToParse) & """:") > 0 Then
            tmpItm = Split(strItm, Trim(vToParse) & """:")(1)
            tmpItm = Split(tmpItm, ",""")(0)
            If Left(tmpItm, 1) = """" Then tmpItm = Right(tmpItm, Len(tmpItm) - 1)
            If Right(tmpItm, 1) = """" Then tmpItm = Left(tmpItm, Len(tmpItm) - 1)
            tmpItm = Replace(tmpItm, "< ", "")
            If Not IsMissing(strToRemove) Then
                For Each vToRemove In vaToRemove
                    tmpItm = Replace(tmpItm, Trim(vToRemove), "")
                Next
            End If
            vaItm(j) = CStr(tmpItm)
        End If
        j = j + 1
    Next
    On Error GoTo 0
    dicItm.Add i, Array(vaItm, 1)
Next

'----------------------------------------------------
'Dictionary -> �迭 ��ȯ
'----------------------------------------------------
r = dicItm.Count
c = UBound(vaToParse) + 1
If Not IsMissing(strID) Then c = c + 1

If r = 0 Then ParseJSON = ""

vaItems = dicItm.Items

ReDim vaReturn(1 To r, 1 To c)

On Error Resume Next
For i = 0 To r - 1
    For j = 0 To c - 1
        tmpItm = vaItems(i)(0)(j)
        If IsNumeric(tmpItm) And Left(tmpItm, 1) <> 0 Then vaReturn(i + 1, j + 1) = CDbl(tmpItm) Else vaReturn(i + 1, j + 1) = tmpItm
    Next
Next
On Error GoTo 0

'----------------------------------------------------
'����� ����
'----------------------------------------------------
ParseJSON = vaReturn

End Function
