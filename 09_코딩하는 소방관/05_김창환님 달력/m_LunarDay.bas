Attribute VB_Name = "m_LunarDay"
Option Explicit

Type Julgi12   '12�߱� ����
  KName As String
  MonNumber As Byte
  Ref_Day As Double
  Longitude As Double
  RealDay As Double
End Type

Type LunarDay  '���� ���
  StartDay As Double
  MonLength As Integer
  Junggi As Boolean ' Byte
  MonName As Byte
  LYear As Integer
End Type

Const MoonMonth As Double = 29.5305882
Const MoonDay As Double = 12.190749387105
Const OneYear As Double = 365.24219
Const Oneday As Double = 360 / OneYear

Dim Junggi(15) As Julgi12
Public LSTable(25) As LunarDay

'����->���
'���͸� �����̸� ���� ��ȯ(��, ū ��, ���� �� �Ǵ��� ����. ������� 12���� ���� ��-29�ϱ��� ����-�ε� 12�� 30���� �־ ������ ��ȯ�ϰ�
'��¥�� �����. �� ��� ����� 12�� 29���� ������, �� ���� 1��1���� ��¥�� ��.
Public Function InvLuniSolarCal(ByVal LunarYear As Integer, ByVal LunarMon As Byte, ByVal LunarDay As Byte, ByVal IsLeap As Boolean, ByVal TZone As Double, _
                                               ByVal MeanSun As Boolean, ByVal MeanMoon As Boolean, ByVal Jinsak As Boolean, JD As Double) As Boolean
  Dim iJD As Double, iLY As Integer, iLM As Byte, iLD As Byte, iLM2 As Single, iLeap As Boolean, i As Integer, D1 As Double, lm2 As Single, IsValid As Boolean
  Dim i2 As Integer
  
  IsValid = False  '�ϴ� ���͸� ������ ����
  
  'step 1. �ʱ� ����ġ ���
  iJD = JULIANDAY(CDbl(LunarYear), CDbl(LunarMon), CDbl(LunarDay), 12, 0)
  iJD = iJD + 25: lm2 = LunarMon
  If IsLeap Then iJD = iJD + 25: lm2 = lm2 + 0.5 '�����̸� +50��, �ƴϸ� +25��
  '������ �׻� ���� �̸��� ��� �ڿ� ���Ƿ� ������ ��� �� �̸��� ��޺��� �ణ ū ���� ����(���⿡���� ���̸�+0.5)
  
  If LunarYear < 1582 Then iJD = iJD + Int(0.0078 * (1582 - LunarYear)) '�����콺�� ������
  
  'step 2. �ʱ� ����ġ�� �������� ���� ���
  i = 0  '���ѷ��� ������
  Do
    LuniSolarCal iJD, iLY, iLM, iLD, iLeap, TZone, MeanSun, MeanMoon, Jinsak
    
    If iLY > LunarYear Then  '���� ������ ���� �Է°����� ũ�� 30�� ���������� �ٽ� ���
      iJD = iJD - 30
    ElseIf iLY < LunarYear Then  '���� ������ ���� �Է°����� ������ 30�� ���İ����� �ٽ� ���
      iJD = iJD + 30
    ElseIf iLY = LunarYear Then  '������ ������
      i2 = 0  '���ѷ��� ������
      Do
        LuniSolarCal iJD, iLY, iLM, iLD, iLeap, TZone, MeanSun, MeanMoon, Jinsak
        iLM2 = iLM + IIf(iLeap, 0.5, 0)  '������ ��� �׻� ���� �̸��� ��� �ڿ� ���Ƿ� ���̸��� ��޺��� �ణ ū ���� ����(���⿡���� ���̸�+0.5)
        
        If iLM2 > lm2 Or iLY > LunarYear Then '���� ��(�Ǵ� ����)�� ���� �Է°����� ũ�� 10�� ���������� �ٽ� ���
          iJD = iJD - 10
        ElseIf iLM2 < lm2 Or iLY < LunarYear Then    '���� ��(�Ǵ� ����)�� ���� �Է°����� ������ 10�� ���İ����� �ٽ� ���
          iJD = iJD + 10
        ElseIf iLM2 = lm2 Then  '���� ����� ã������
          D1 = iJD - iLD  '���� �ش� ���� 0�� ��¥
          IsValid = True  '�� ã�����Ƿ� ��ȿ���� ������ ��ħ
        End If
        i2 = i2 + 1
      Loop Until iLM2 = lm2 Or i2 > 15
    End If
    i = i + 1
  Loop Until iLY = LunarYear Or i > 10
  
  JD = D1 + LunarDay  '�ش� ���� 0�Ͽ� �Է��� ���� ��¥�� ����.
  InvLuniSolarCal = IsValid
End Function

'���->����
Public Sub LuniSolarCal(ByVal JD As Double, LunarYear As Integer, LunarMon As Byte, LunarDay As Byte, IsLeap As Boolean, _
                                   ByVal TZone As Double, ByVal MeanSun As Boolean, ByVal MeanMoon As Boolean, ByVal Jinsak As Boolean)
  Dim jd0 As Double, dYear As Double, yJD0 As Double, bf As Double, B As Integer, M As Integer
  Dim i As Integer, j As Integer, k As Integer, LD(25) As LunarDay, SD(25) As Double
  Dim PreWinter As Double, ThisWinter As Double, Count1 As Integer, idx1 As Integer, idx2 As Integer
  Dim LeapType As Byte, Leap13 As Boolean, fMON As Integer, a As Integer, LCount As Integer
  
  jd0 = GetJD0(JD) + 0.5
  dYear = InvJDYear(jd0)
  
  '�Է����� õ�������Ϻ��� ���̸� ���� �طΰ��
  If MeanSun Then  '����
    If jd0 < Pyunggi(dYear, 270, -13, TZone) Then dYear = dYear - 1
    If jd0 > Pyunggi(dYear, 270, 355, TZone) Then dYear = dYear + 1
  Else  '�����
    If jd0 < cJunggi(dYear, 270, -13, TZone) Then dYear = dYear - 1
    If jd0 > cJunggi(dYear, 270, 355, TZone) Then dYear = dYear + 1
  End If
  
  yJD0 = JULIANDAY(dYear, 1, 1, 12, 0)
  Call Set24Julgi: Call CalcJulGi(yJD0, TZone, MeanSun)   '�Ա� �ð� �� ����ð� ��� �κ���
  
  j = 0: SD(0) = 0: k = 0
  Do
    If MeanMoon Then  '����
      bf = GetJD0(GetMeanMoon(yJD0 - 96 + j * 28, TZone)) + 0.5
    Else  '�����, �����
      bf = GetMoon(yJD0 - 96 + j * 28, 0, TZone)
      If Jinsak Then  ' �����
        If bf - GetJD0(bf) < 0.75 Then bf = GetJD0(bf) + 0.5 Else bf = GetJD0(bf) + 1.5
      Else '�����
        bf = GetJD0(bf) + 0.5
      End If
    End If
    
    If bf >= Junggi(0).RealDay Then
      If k > 0 Then
        If bf > SD(k - 1) Then SD(k) = bf: k = k + 1
      Else
        SD(0) = bf: k = 1
      End If
    End If
    j = j + 1
  Loop Until bf > yJD0 + 427
  j = k: k = j - 1
  
  For i = 0 To 25 '���� �ʱ�ȭ
    LD(i).StartDay = 0
    LD(i).StartDay = SD(i)
    LD(i).MonName = 100
  Next i
  PreWinter = Junggi(1).RealDay
  ThisWinter = Junggi(13).RealDay
  
  idx1 = 0: idx2 = 0
  For i = 0 To 24 '�Ա�ð��� ������� �������� �� �̸��� ���̰� ���߿��� �Ǻ�
    LD(i).Junggi = False
    For j = 0 To 15
      If LD(i + 1).StartDay > Junggi(j).RealDay And Junggi(j).RealDay >= LD(i).StartDay Then
        LD(i).Junggi = True
        If LD(i).MonName = 100 Then
          LD(i).MonName = Junggi(j).MonNumber
        ElseIf Junggi(j).MonNumber = 11 Or LD(i).MonName = 11 Then
          LD(i).MonName = 11 '�������� ������ 11��
        End If
      End If
    Next j
  Next i
  
  Count1 = 0 '���� �����ް� �̹� ������ ������ ����� �� ����
  For i = 0 To k - 1
    If PreWinter < LD(i).StartDay And LD(i).StartDay <= ThisWinter Then Count1 = Count1 + 1
    If PreWinter < LD(i + 1).StartDay And LD(i).StartDay <= PreWinter Then idx1 = i
    If ThisWinter < LD(i + 1).StartDay And LD(i).StartDay <= ThisWinter Then idx2 = i
  Next i
  
  '������� ���� ����(2009.2.19)================
  '���⼭���� ġ�� �� ����ȣ �ű�� ����
  If Count1 = 12 Then '�� ��� õ�������� �ݳ��� ���� ���̿��� ��� 12����. ���߿��� �ֵ���� �� ���̴� ������ ���.
    LeapType = 4
  Else 'count1=13�� ��. ���� �˻� �ʿ�
    If LD(idx1 + 1).Junggi = True And LD(idx1 + 2).Junggi = True Then LeapType = 1 '���� �������� ���߿��̰� �� �����޵� ���߿�. ��12���� ��1��.
                                                                                   '�� ��� 1��~11�� ���̿� ����
    If LD(idx1 + 1).Junggi = False And LD(idx1 + 2).Junggi = True Then LeapType = 2 '���� �������� ���߿��̰� �� �������� ���߿�. ��11���� ��12��
    If LD(idx1 + 1).Junggi = True And LD(idx1 + 2).Junggi = False Then LeapType = 3 '���� �������� ���߿��̰� �� ���� ���� ���߿�. ��12���� ��12��.
    
    Leap13 = False '������ ��ġ ����
    For i = (idx1 + 3) To (idx2 - 1)  '(õ��������+3��)°���� �ݳ� ���������� �˻�
      If LD(i).Junggi = False Then Leap13 = True 'LeapType=1�� ��1��~11���߿� ��� ������ �ִ��� Ž��. Leaptype= 2�� 3�� ���� ���ǹ�.
                                                 '������ �̹� ���� 11�� �Ǵ� 12���� �پ� �����Ƿ� �� ���Ĵ� ��� ��޷� ���(Leap13 = False)
    Next i
    If LeapType = 2 Or LeapType = 3 Then Leap13 = False
  End If
  '�� �� �������� ���� ������ ������� 13������ �Ǵ� ���� ����. �׷��Ƿ� count1=13�� �Ǵ� ���� �� �ش� �ݵ�� count1=12��.
  '�� ���� �ص� �ݵ�� count1=12��.
  '��, ������ �Ǳ� ���� �޿� ������ ���� �� ����. ���� ������ 12���� ������� ��� ����� �Ǿ�� ��.
  
  LCount = 0
  Select Case LeapType
   Case 1, 4 '������ LD(idx1+3)�� LD(idx2-1) ���̿� ����[(õ��������+3��)°���� �ݳ� ���� ����]
     LD(idx1 + 1).MonName = 12: LD(idx1 + 1).LYear = dYear - 1: LD(idx1 + 1).Junggi = True
     LD(idx1 + 2).MonName = 1: LD(idx1 + 2).LYear = dYear: LD(idx1 + 2).Junggi = True
     
   Case 2
      LD(idx1 + 1).MonName = 11: LD(idx1 + 1).LYear = dYear - 1: LD(idx1 + 1).Junggi = False
      LD(idx1 + 2).MonName = 12: LD(idx1 + 2).LYear = dYear - 1
     
   Case 3
     LD(idx1 + 1).MonName = 12: LD(idx1 + 1).LYear = dYear - 1
     LD(idx1 + 2).MonName = 12: LD(idx1 + 2).LYear = dYear - 1: LD(idx1 + 2).Junggi = False
  
  End Select
  LD(idx1).MonName = 11: LD(idx1).LYear = dYear - 1
  
  '������� ���캸�� ��� ��쿡�� ���� �� 1���� ���޷� �Ǵ� ���� ����.
  '�� ������ õ�������� 1���� ���̷κ��� ������ 1���� ��ġ�� �����ϴ� ������.
  '���� (idx1+3)���� ���� ġ���ϸ� ��.
  
  'LeapType=1�� ������ ��1��(�Ǵ� ��2��)���� ��11��(������)����(���̿� ���� �ϳ� ����)
  'LeapType=2,3�� ������ ��1������ �� 11������(���̴� ��� ���)
  'LeapType=4�� ������ ��2������ �� 11������(���̴� ��� ���)
  
  fMON = 1: a = 0
  If LeapType = 4 Then a = 1
  For i = idx1 + 3 To idx2
    LD(i).LYear = dYear '������ ������ ������.
    
    If LeapType = 1 Then
      If LD(i).Junggi = True Or LCount > 0 Then
        '���߿��� �ƴϰų� �ι�° ���߿��̸�, ��޷� ó��.
        LD(i).Junggi = True
        a = a + 1
      Else
        '�����̸� ���� �� �� ����.
        LCount = 1
        
      End If
      LD(i).MonName = fMON + a
    
    Else
      LD(i).MonName = fMON + a
      LD(i).Junggi = True  '������ ���.
      a = a + 1
    End If
  Next i  '������� ����(2009.2.19)================
  
  For i = 0 To 24 '�� ���� ���
    If Abs(LD(i + 1).StartDay - LD(i).StartDay) < 31 Then
      LD(i).MonLength = LD(i + 1).StartDay - LD(i).StartDay
    End If
  Next i
  
  '���� ���
  For i = 0 To k - 1
    If jd0 >= LD(i).StartDay And jd0 < LD(i + 1).StartDay Then
      LunarYear = LD(i).LYear
      LunarMon = LD(i).MonName
      LunarDay = jd0 - LD(i).StartDay + 1
      IsLeap = Not LD(i).Junggi
      Exit For
    End If
  Next i
End Sub

'�����Լ���
Sub CalcJulGi(ByVal JDt As Double, ByVal TZone As Double, ByVal MeanSun As Boolean)
  Dim i As Integer, nYear As Double
  
  nYear = InvJDYear(JDt)
  
  For i = 0 To 15
    If MeanSun Then
      Junggi(i).RealDay = GetJD0(Pyunggi(nYear, Junggi(i).Longitude, Junggi(i).Ref_Day, TZone)) + 0.5
    Else
      Junggi(i).RealDay = GetJD0(cJunggi(nYear, Junggi(i).Longitude, Junggi(i).Ref_Day, TZone)) + 0.5
    End If
  Next i
End Sub

'����� LST�� ����.  TDT ��길 ��Ȯ�ϴٸ� ���е��� -2000~3000�� ���̿��� 1�� ����
Function GetMoon(ByVal cJD As Double, ByVal LonMoon As Double, ByVal TZone As Double) As Double
  Dim k As Long, PosMoon As TPlanetData, PosSun As TPlanetData
  Dim LamSun As Double, dt As Double, dLam As Double, tJD As Double
  Dim LamMoon As Double, MAge As Double, i As Long, fFlag As Boolean
  Dim tTDT As Double
  
  dt = 0: i = 0: fFlag = True
  
  tJD = cJD
  tTDT = JDtoTDT(tJD)
  
start:
  PosSun.JD = tTDT
  PosSun.ipla = 3
  PosMoon.JD = tTDT
  PosMoon.ipla = 11
  k = Plan404(PosSun)
  k = Plan404(PosMoon)
  LamSun = Rev(PosSun.l * RadtoDeg + 180)
  LamMoon = Rev(PosMoon.l * RadtoDeg)
  MAge = Rev(LamMoon - LamSun)
  dLam = AngDistLon(MAge, LonMoon)
  
  If (LonMoon > 357 Or LonMoon < 3) And fFlag Then
    tJD = cJD - MAge / MoonDay
    tTDT = JDtoTDT(tJD)
    dt = 0
    fFlag = False
    GoTo start
  End If
  
  Do
    dt = dLam / MoonDay
    
    If LonMoon > 357 Or LonMoon < 3 Then
      If MAge > 180 Then MAge = MAge - 360
    End If
    
    If LonMoon > MAge Then
      tJD = tJD + dt
    Else
      tJD = tJD - dt
    End If
    
    tTDT = JDtoTDT(tJD)
    PosSun.JD = tTDT
    PosSun.ipla = 3
    PosMoon.JD = tTDT '���� ����, �嵿 ������ ���ص� ��(�ؿ� ���� ��� ��ġ�� �˸� �ǹǷ�)
    PosMoon.ipla = 11
    k = Plan404(PosSun)
    k = Plan404(PosMoon)
    LamSun = Rev(PosSun.l * RadtoDeg + 180 - 0.005691611 / PosSun.R)
    LamMoon = Rev(PosMoon.l * RadtoDeg)
    MAge = Rev(LamMoon - LamSun)
    dLam = AngDistLon(MAge, LonMoon)
    i = i + 1
  Loop Until (dLam / MoonDay * 86400) < 0.1 Or i > 50  '�ð� ����: 0.1�� �̳�
  
  GetMoon = tJD + TZone / 24 ': Debug.Print i
End Function

Function cJunggi(ByVal cYear As Double, ByVal LonSun As Double, ByVal RefDay As Double, ByVal TZone As Double) As Double
  Dim JDyear As Double, aDay As Double, PosSun As TPlanetData, k As Long
  Dim LamSun As Double, dt As Double, dLam As Double, tJD As Double, i As Long
  Dim dl As Double, de As Double
  
  dt = 0: i = 0
  If cYear < 1582 Then RefDay = RefDay + Int(0.0078 * (1582 - cYear))
  JDyear = JULIANDAY(cYear, 1, 0, 0, 0, 0)
  
  tJD = JDyear + RefDay
  PosSun.JD = JDtoTDT(tJD)
  PosSun.ipla = 3
  k = Plan404(PosSun)
  LamSun = Rev(PosSun.l * RadtoDeg + 180 - 0.005691611 / PosSun.R)
  dLam = AngDistLon(LamSun, LonSun)
  
  Do
    dt = dLam / Oneday
    
    If LonSun > 357 Or LonSun < 3 Then
      If LamSun > 180 Then LamSun = LamSun - 360
    End If
    
    If LonSun > LamSun Then
      tJD = tJD + dt
    Else
      tJD = tJD - dt
    End If
    
    PosSun.JD = JDtoTDT(tJD)
    PosSun.ipla = 3
    k = Plan404(PosSun)
    Nutation PosSun.JD, dl, de
    LamSun = Rev(PosSun.l * RadtoDeg + 180 + dl / 3600 - 0.005691611 / PosSun.R)  '�嵿 ����, 5.69161111111111E-03�� ������
    dLam = AngDistLon(LamSun, LonSun)
    i = i + 1
  Loop Until (dLam / Oneday * 86400) < 0.1 Or i > 50 '�ð� ����: 0.1�� �̳�
  cJunggi = tJD + TZone / 24
End Function

'24���� �Է�
Public Sub Set24Julgi()
  Junggi(0).KName = "�Ҽ�": Junggi(0).MonNumber = 10: Junggi(0).Longitude = 240: Junggi(0).Ref_Day = -43 '11
  Junggi(1).KName = "����": Junggi(1).MonNumber = 11: Junggi(1).Longitude = 270: Junggi(1).Ref_Day = -13     '12
  Junggi(2).KName = "����": Junggi(2).MonNumber = 12: Junggi(2).Longitude = 300: Junggi(2).Ref_Day = 20     '1
  Junggi(3).KName = "���": Junggi(3).MonNumber = 1: Junggi(3).Longitude = 330: Junggi(3).Ref_Day = 50     '2
  Junggi(4).KName = "���": Junggi(4).MonNumber = 2: Junggi(4).Longitude = 0: Junggi(4).Ref_Day = 80     '3
  Junggi(5).KName = "���": Junggi(5).MonNumber = 3: Junggi(5).Longitude = 30: Junggi(5).Ref_Day = 110     '4
  Junggi(6).KName = "�Ҹ�": Junggi(6).MonNumber = 4: Junggi(6).Longitude = 60: Junggi(6).Ref_Day = 140     '5
  Junggi(7).KName = "����": Junggi(7).MonNumber = 5: Junggi(7).Longitude = 90: Junggi(7).Ref_Day = 170     '6
  Junggi(8).KName = "�뼭": Junggi(8).MonNumber = 6: Junggi(8).Longitude = 120: Junggi(8).Ref_Day = 200      '7
  Junggi(9).KName = "ó��": Junggi(9).MonNumber = 7: Junggi(9).Longitude = 150: Junggi(9).Ref_Day = 230      '8
  Junggi(10).KName = "�ߺ�": Junggi(10).MonNumber = 8: Junggi(10).Longitude = 180: Junggi(10).Ref_Day = 260    '9
  Junggi(11).KName = "��": Junggi(11).MonNumber = 9: Junggi(11).Longitude = 210: Junggi(11).Ref_Day = 290     '10
  Junggi(12).KName = "�Ҽ�": Junggi(12).MonNumber = 10: Junggi(12).Longitude = 240: Junggi(12).Ref_Day = 320      '11
  Junggi(13).KName = "����": Junggi(13).MonNumber = 11: Junggi(13).Longitude = 270: Junggi(13).Ref_Day = 350      '12
  Junggi(14).KName = "����": Junggi(14).MonNumber = 12: Junggi(14).Longitude = 300: Junggi(14).Ref_Day = 385     '1
  Junggi(15).KName = "���": Junggi(15).MonNumber = 1: Junggi(15).Longitude = 330: Junggi(15).Ref_Day = 415     '2
End Sub


'High Speed
'���->����
Public Sub LSTBL(ByVal JD As Double, ByVal TZone As Double, ByVal MeanSun As Boolean, ByVal MeanMoon As Boolean, ByVal Jinsak As Boolean)
  Dim jd0 As Double, dYear As Double, yJD0 As Double, bf As Double, B As Integer, M As Integer
  Dim i As Integer, j As Integer, k As Integer, LD(25) As LunarDay, SD(25) As Double
  Dim PreWinter As Double, ThisWinter As Double, Count1 As Integer, idx1 As Integer, idx2 As Integer
  Dim LeapType As Byte, Leap13 As Boolean, fMON As Integer, a As Integer, LCount As Integer
  
  jd0 = GetJD0(JD) + 0.5
  dYear = InvJDYear(jd0)
  yJD0 = JULIANDAY(dYear, 1, 1, 12, 0)
  
  Call Set24Julgi: Call CalcJulGi(yJD0, TZone, MeanSun)   '�Ա� �ð� �� ����ð� ��� �κ���
  
  j = 0: SD(0) = 0: k = 0
  Do
    If MeanMoon Then  '����
      bf = GetJD0(GetMeanMoon(yJD0 - 96 + j * 28, TZone)) + 0.5
    Else  '�����, �����
      bf = GetMoon(yJD0 - 96 + j * 28, 0, TZone)
      If Jinsak Then  ' �����
        If bf - GetJD0(bf) < 0.75 Then bf = GetJD0(bf) + 0.5 Else bf = GetJD0(bf) + 1.5
      Else '�����
        bf = GetJD0(bf) + 0.5
      End If
    End If
    
    If bf >= Junggi(0).RealDay Then
      If k > 0 Then
        If bf > SD(k - 1) Then SD(k) = bf: k = k + 1
      Else
        SD(0) = bf: k = 1
      End If
    End If
    j = j + 1
  Loop Until bf > yJD0 + 427
  j = k: k = j - 1
  
  For i = 0 To 25 '���� �ʱ�ȭ
    LD(i).StartDay = 0
    LD(i).StartDay = SD(i)
    LD(i).MonName = 100
  Next i
  PreWinter = Junggi(1).RealDay
  ThisWinter = Junggi(13).RealDay
  
  idx1 = 0: idx2 = 0
  For i = 0 To 24  '�Ա� �ð��� ������� �������� �� �̸��� ���̰� ���߿��� �Ǻ�
    LD(i).Junggi = False
    For j = 0 To 15
      If LD(i + 1).StartDay > Junggi(j).RealDay And Junggi(j).RealDay >= LD(i).StartDay Then
        LD(i).Junggi = True
        If LD(i).MonName = 100 Then
          LD(i).MonName = Junggi(j).MonNumber
        ElseIf Junggi(j).MonNumber = 11 Or LD(i).MonName = 11 Then
            LD(i).MonName = 11 '�������� ������ 11��
        End If
      End If
    Next j
  Next i
  
  Count1 = 0  '���� �����ް� �̹� ������ ������ ����� �� ����
  For i = 0 To k - 1
    If PreWinter < LD(i).StartDay And LD(i).StartDay <= ThisWinter Then Count1 = Count1 + 1
    If PreWinter < LD(i + 1).StartDay And LD(i).StartDay <= PreWinter Then idx1 = i
    If ThisWinter < LD(i + 1).StartDay And LD(i).StartDay <= ThisWinter Then idx2 = i
  Next i
  
  '������� ���� ����(2009.2.19)================
  '���⼭���� ġ�� �� ����ȣ �ű�� ����
  If Count1 = 12 Then '�� ��� õ�������� �ݳ��� ���� ���̿��� ��� 12����. ���߿��� �ֵ���� �� ���̴� ������ ���.
    LeapType = 4
  Else 'count1=13�� ��. ���� �˻� �ʿ�
    If LD(idx1 + 1).Junggi = True And LD(idx1 + 2).Junggi = True Then LeapType = 1 '���� �������� ���߿��̰� �� �����޵� ���߿�. ��12���� ��1��.
                                                                                   '�� ��� 1��~11�� ���̿� ����
    If LD(idx1 + 1).Junggi = False And LD(idx1 + 2).Junggi = True Then LeapType = 2 '���� �������� ���߿��̰� �� �������� ���߿�. ��11���� ��12��
    If LD(idx1 + 1).Junggi = True And LD(idx1 + 2).Junggi = False Then LeapType = 3 '���� �������� ���߿��̰� �� ���� ���� ���߿�. ��12���� ��12��.
    
    Leap13 = False '������ ��ġ ����
    For i = (idx1 + 3) To (idx2 - 1)  '(õ��������+3��)°���� �ݳ� ���������� �˻�
      If LD(i).Junggi = False Then Leap13 = True 'LeapType=1�� ��1��~11���߿� ��� ������ �ִ��� Ž��. Leaptype= 2�� 3�� ���� ���ǹ�.
                                                 '������ �̹� ���� 11�� �Ǵ� 12���� �پ� �����Ƿ� �� ���Ĵ� ��� ��޷� ���(Leap13 = False)
    Next i
    If LeapType = 2 Or LeapType = 3 Then Leap13 = False
  End If
  '�� �� �������� ���� ������ ������� 13������ �Ǵ� ���� ����. �׷��Ƿ� count1=13�� �Ǵ� ���� �� �ش� �ݵ�� count1=12��.
  '�� ���� �ص� �ݵ�� count1=12��.
  '��, ������ �Ǳ� ���� �޿� ������ ���� �� ����. ���� ������ 12���� ������� ��� ����� �Ǿ�� ��.
  
  LCount = 0
  Select Case LeapType
   Case 1, 4 '������ LD(idx1+3)�� LD(idx2-1) ���̿� ����[(õ��������+3��)°���� �ݳ� ���� ����]
     LD(idx1 + 1).MonName = 12: LD(idx1 + 1).LYear = dYear - 1: LD(idx1 + 1).Junggi = True
     LD(idx1 + 2).MonName = 1: LD(idx1 + 2).LYear = dYear: LD(idx1 + 2).Junggi = True
     
   Case 2
      LD(idx1 + 1).MonName = 11: LD(idx1 + 1).LYear = dYear - 1: LD(idx1 + 1).Junggi = False
      LD(idx1 + 2).MonName = 12: LD(idx1 + 2).LYear = dYear - 1
     
   Case 3
     LD(idx1 + 1).MonName = 12: LD(idx1 + 1).LYear = dYear - 1
     LD(idx1 + 2).MonName = 12: LD(idx1 + 2).LYear = dYear - 1: LD(idx1 + 2).Junggi = False
  
  End Select
  LD(idx1).MonName = 11: LD(idx1).LYear = dYear - 1
  
  '������� ���캸�� ��� ��쿡�� ���� �� 1���� ���޷� �Ǵ� ���� ����.
  '�� ������ õ�������� 1���� ���̷κ��� ������ 1���� ��ġ�� �����ϴ� ������.
  '���� (idx1+3)���� ���� ġ���ϸ� ��.
  
  'LeapType=1�� ������ ��1��(�Ǵ� ��2��)���� ��11��(������)����(���̿� ���� �ϳ� ����)
  'LeapType=2,3�� ������ ��1������ �� 11������(���̴� ��� ���)
  'LeapType=4�� ������ ��2������ �� 11������(���̴� ��� ���)
  
  fMON = 1: a = 0
  If LeapType = 4 Then a = 1
  For i = idx1 + 3 To idx2
    LD(i).LYear = dYear '������ ������ ������.
    
    If LeapType = 1 Then
      If LD(i).Junggi = True Or LCount > 0 Then
        '���߿��� �ƴϰų� �ι�° ���߿��̸�, ��޷� ó��.
        LD(i).Junggi = True
        a = a + 1
      Else
        '�����̸� ���� �� �� ����.
        LCount = 1
        
      End If
      LD(i).MonName = fMON + a
    
    Else
      LD(i).MonName = fMON + a
      LD(i).Junggi = True  '������ ���.
      a = a + 1
    End If
  Next i  '������� ����(2009.2.19)================
  
  For i = 0 To 24 '�� ���� ���
    If Abs(LD(i + 1).StartDay - LD(i).StartDay) < 31 Then
      LD(i).MonLength = LD(i + 1).StartDay - LD(i).StartDay
    End If
  Next i
  
  '���� ��� �κ�(��� ������ ������ ��Ȯ�� ���Ǵ� ������ ����)
  B = 0
  For i = idx1 To idx2
    With LSTable(B)
      .Junggi = LD(i).Junggi
      .LYear = LD(i).LYear
      .MonLength = LD(i).MonLength
      .MonName = LD(i).MonName
      .StartDay = LD(i).StartDay
    End With
    B = B + 1
  Next i
End Sub

Sub LSTBL2(ByVal JD As Double, ByVal TZone As Double, ByVal MeanSun As Boolean, ByVal MeanMoon As Boolean, ByVal Jinsak As Boolean)
  Dim tLST(75) As LunarDay, i As Integer, dYear As Double, yJD As Double
  Dim N As Integer, j As Integer, a As Boolean, B As Boolean
  
  '3��ġ ǥ ����
  dYear = InvJDYear(JD) - 1: N = 0
  For i = 0 To 2
    yJD = JULIANDAY(dYear + i, 1, 1, 12, 0)
    ClearLSTBL
'    If AutoConfig = True Then Call AutoChoose(CInt(dYear + i))
    LSTBL yJD, TZone, UseMeanSun, UseMeanMoon, UseJinsak   '�� �Լ��� ����ϴ� ������ lstable�� 0~13 �����̹Ƿ�
    For N = 0 To 25 '�� �������� ǥ�� �Ϻΰ� ���ĵ� ��� ����
      tLST(N + 25 * i).Junggi = LSTable(N).Junggi
      tLST(N + 25 * i).LYear = LSTable(N).LYear
      tLST(N + 25 * i).MonLength = LSTable(N).MonLength
      tLST(N + 25 * i).MonName = LSTable(N).MonName
      tLST(N + 25 * i).StartDay = LSTable(N).StartDay
    Next N
  Next i
  
  '��¥ ������ ���� ����
  For i = 0 To 74
    yJD = tLST(i).StartDay
    N = i
    For j = i + 1 To 75
      If tLST(j).StartDay < yJD Then
        yJD = tLST(j).StartDay
        N = j
      End If
    Next j
    SwapLD tLST(i), tLST(N)
  Next i
  
  '�ݺ��װ� ���ʿ��� �κ� ����
  ClearLSTBL
  N = 0: j = CInt(dYear) + 1
  For i = 0 To 74
    a = (tLST(i).LYear = j - 1) And ((tLST(i).MonName > 6) And (tLST(i).MonName < 13))
    a = a Or (tLST(i).LYear = j) And ((tLST(i).MonName > 0) And (tLST(i).MonName < 13))
    a = a Or (tLST(i).LYear = j + 1) And ((tLST(i).MonName < 4) And (tLST(i).MonName > 0))
    B = tLST(i).StartDay <> tLST(i + 1).StartDay
    
    If a And B Then
      LSTable(N).Junggi = tLST(i).Junggi
      LSTable(N).LYear = tLST(i).LYear
      LSTable(N).MonLength = tLST(i).MonLength
      LSTable(N).MonName = tLST(i).MonName
      LSTable(N).StartDay = tLST(i).StartDay
'      Debug.Print n; LSTable(n).LYear; LSTable(n).MonName; LSTable(n).Junggi, LSTable(n).StartDay
      N = N + 1
    End If
  Next i
End Sub

Private Sub SwapLD(Var1 As LunarDay, Var2 As LunarDay)
  Dim Te As LunarDay
  
  Te.Junggi = Var1.Junggi: Te.LYear = Var1.LYear
  Te.MonLength = Var1.MonLength: Te.MonName = Var1.MonName
  Te.StartDay = Var1.StartDay
  
  Var1.Junggi = Var2.Junggi: Var1.LYear = Var2.LYear
  Var1.MonLength = Var2.MonLength: Var1.MonName = Var2.MonName
  Var1.StartDay = Var2.StartDay
  
  Var2.Junggi = Te.Junggi: Var2.LYear = Te.LYear
  Var2.MonLength = Te.MonLength: Var2.MonName = Te.MonName
  Var2.StartDay = Te.StartDay
End Sub

Private Sub ClearLSTBL()
 Dim i As Integer
 
 For i = 0 To 25
   With LSTable(i)
     .Junggi = False
     .LYear = 0
     .MonLength = 0
     .MonName = 0
     .StartDay = 0
   End With
 Next i
End Sub

Sub FindTBL(ByVal JD As Double, LunarYear As Integer, LunarMon As Byte, LunarDay As Byte, IsLeap As Boolean)
  Dim i As Integer, jd0 As Double
  
  i = 0
  jd0 = GetJD0(JD) + 0.5
  Do While LSTable(i).StartDay <= jd0 And i < 25
    If jd0 >= LSTable(i).StartDay And jd0 < LSTable(i + 1).StartDay Then
      LunarYear = LSTable(i).LYear
      LunarMon = LSTable(i).MonName
      LunarDay = jd0 - LSTable(i).StartDay + 1
      IsLeap = Not LSTable(i).Junggi
    End If
    
    i = i + 1
  Loop
End Sub

Function FindTBLInv(ByVal LunarYear As Integer, ByVal LunarMon As Byte, ByVal LunarDay As Byte, ByVal IsLeap As Boolean, JD As Double) As Boolean
  Dim i As Integer, k As Double, a As Boolean
 
  i = 0: a = False: JD = 0
  Do While i < 25 And a = False
    If LunarYear = LSTable(i).LYear Then
      If LunarMon = LSTable(i).MonName And IsLeap = Not LSTable(i).Junggi Then
        k = LSTable(i).StartDay: a = True
        JD = k + LunarDay - 1
      End If
    End If
    i = i + 1
  Loop
  
  FindTBLInv = a
End Function

