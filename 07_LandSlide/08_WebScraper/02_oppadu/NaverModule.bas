Attribute VB_Name = "NaverModule"
Option Explicit


Sub NaverLandCrawl()

'##########################################################
' ���� ����
Dim htmlResult As Object
Dim strResult As String
Dim url As String
Dim city As String


'##########################################################
'1. ���̹� �ε��� ���� ������ -> ���� �˻�
'                                               -> ���� �Ź� ����� �޾ƿ��� ���� ��/�浵, �� �� ���� ����
'[            https://meyerweb.com/eric/tools/dencoder/                ]  '-> URL ���ڵ�/���ڵ�
'[            http://json.parser.online.fr/                                             ]  '-> JSON �ļ�

city = Sheet1.Range("D5")
url = "https://m.land.naver.com/search/result/" & city

'Debug.Print url

Set htmlResult = GetHttp(url)
strResult = htmlResult.body.innerHTML

strResult = Splitter(strResult, "filter: {", "},")

'Debug.Print strResult
'ExportText strResult


'----------------------------------------
' �Ź����� ��) ����Ʈ: APT, ���� : VL, ���ǽ���: OPST, ...������ ����
' �ǹ����� A1, B1, ... ������ ��������
'----------------------------------------



'---------------- �ش� ����(����) ��/�浵 + �� �� ���� ����
Dim lat As String: lat = Splitter(strResult, "lat: '", "',")
Dim lon As String: lon = Splitter(strResult, "lon: '", "',")
Dim z As String: z = Splitter(strResult, "z: '", "',")
Dim cortarNo As String: cortarNo = Splitter(strResult, "cortarNo: '", "',")
Dim searchType As String: searchType = "APT"
Dim buildingType As String: buildingType = "A1:B1:B2"

'#########################################################
' 2. ���̹� �ε��� ���� ������ -> �޾ƿ� ��/�浵, �� �� ������ �˻�
'                                                -> �ش� ���� ����Ź� ����� �� ��/�浵 ����


url = "https://m.land.naver.com/cluster/clusterList?view=actl&cortarNo=" & cortarNo & "&rletTpCd=" & searchType & "&tradTpCd=" & buildingType & _
            "&z=" & z & "&lat=" & lat & "&lon=" & lon & "&addon=COMPLEX&bAddon=COMPLEX&isOnlyIsale=false"

Set htmlResult = Nothing
Set htmlResult = GetHttp(url)
strResult = htmlResult.body.innerHTML

Dim forceCOMPLEX As Boolean
If InStr(1, strResult, "COMPLEX") > 0 Then strResult = Splitter(strResult, "COMPLEX"): forceCOMPLEX = True
Dim v As Variant
v = ParseJSON(strResult, "lgeo,lat,lon,count")

'##########################################################
' 3. ���̹� �ε��� �Ź�������  -> �� �� �浵�� �Ź� ���� �˻�
'                                                -> �� �Ź��� ������ ����
' �Ź��������� �������� 20������ ���.. �׷��� �߰� �۾� �ʿ�!
' Set htmlResult = Nothing
' Set htmlResult = GetHttp(URL)
'                strResult = htmlResult.body.innerHTML
'                vReturn = ParseJSON(strResult, "hscpNm,hscpNo,scpTypeCd,hscpTypeNm,totDongCnt,totHsehCnt,genHsehCnt,useAprvYmd,repImgUrl,dealCnt,leaseCnt,rentCnt," & _
                                             "strmRentCnt,totalAtclCnt,minSpc,maxSpc,dealPrcMin,dealPrcMax,leasePrcMin,leasePrcMax,isalePrcMin,isalePrcMax,isaleNotifSeq,isaleScheLabel,isaleScheLabelPre", city, ",")

Dim i As Long: Dim iPage As Long: Dim j As Long
Dim vReturn As Variant
Dim x As Long: x = GetLastRow(Sheet1) + 1
Dim initR As Long
initR = x

For i = LBound(v, 1) To UBound(v, 1)
    If v(i, 1) <> "" Then
        iPage = Application.WorksheetFunction.RoundUp(v(i, 4) / 20, 0)
        For j = 1 To iPage
            If forceCOMPLEX = False Then
                url = "https://m.land.naver.com/cluster/ajax/articleList?itemId=" & v(i, 1) & "&lgeo=" & v(i, 1) & _
                                "&rletTpCd=" & searchType & "&tradTpCd=" & buildingType & "&z=" & z & "&lat=" & v(i, 2) & "&lon=" & v(i, 3) & "&cortarNo=" & cortarNo & _
                                "&isOnlyIsale=false&sort=readRank&page=" & j
                Debug.Print url
                Set htmlResult = Nothing
                Set htmlResult = GetHttp(url)
                strResult = htmlResult.body.innerHTML
                If InStr(1, strResult, "atclNo") > 0 Then
                    vReturn = ParseJSON(strResult, "atclNm,atclNo,tradTpCd,rletTpNm,totDongCnt_tmp,totHsehCnt_tmp,genHsehCnt_tmp,atclCfmYmd,repImgUrl,tradTpNm,flrInfo,atclTetrDesc," & _
                                    "strmRentCnt_tmp,totalAtclCnt_tmp,spc1,spc2,sameAddrMinPrc,sameAddrMaxPrc,minMviFee,maxMviFee,cpid,cpNm,rltrNm,isaleScheLabel_tmp,isaleScheLabelPre_tmp", city, ",")
                    ArrayToRng Sheet1.Cells(x, 4), vReturn
                    x = x + UBound(vReturn, 1)
                End If
            Else
                url = "https://m.land.naver.com/cluster/ajax/complexList?itemId=" & v(i, 1) & "&lgeo=" & v(i, 1) & _
                        "&rletTpCd=" & searchType & "&tradTpCd=" & buildingType & "&z=" & z & "&lat=" & v(i, 2) & "&lon=" & v(i, 3) & "&cortarNo=" & cortarNo & "&isOnlyIsale=false&sort=readRank&page=" & j
                Set htmlResult = Nothing
                Set htmlResult = GetHttp(url)
                strResult = htmlResult.body.innerHTML
                If InStr(1, strResult, "hscpNo") > 0 Then
                    vReturn = ParseJSON(strResult, "hscpNm,hscpNo,scpTypeCd,hscpTypeNm,totDongCnt,totHsehCnt,genHsehCnt,useAprvYmd,repImgUrl,dealCnt,leaseCnt,rentCnt," & _
                                    "strmRentCnt,totalAtclCnt,minSpc,maxSpc,dealPrcMin,dealPrcMax,leasePrcMin,leasePrcMax,isalePrcMin,isalePrcMax,isaleNotifSeq,isaleScheLabel,isaleScheLabelPre", city, ",")
                    ArrayToRng Sheet1.Cells(x, 4), vReturn
                    x = x + UBound(vReturn, 1)
                End If
            End If
        Next j
    End If
Next i

'#########################################################
' 4. ��ǥ �̹��� ����

'    Dim shpImg As Shape: Dim shpRng As Range
'    For j = initR To Sheet1.Cells(Sheet1.Rows.Count, 4).End(xlUp).Row + 1
'        Set shpRng = Sheet1.Cells(j, 13)
'        If shpRng.Value <> "" Then
'            shpRng.EntireRow.RowHeight = 80
'            InsertWebImage shpRng, "https://landthumb-phinf.pstatic.net" & shpRng.Value
'        Else
'            shpRng.EntireRow.RowHeight = 18
'        End If
'    Next
'

End Sub



