Attribute VB_Name = "m_RiseSet"
Option Explicit

Type tRSTime
  JD As Double  '��¥
  ObjAlt As Double  '��ǥ ��
  Longitude As Double  '�浵
  Latitude As Double  '����
  TZone As Double '�ð���
  
  RiseTime As Double  '�ߴ� �ð�
  SetTime As Double  '���� �ð�
  
  bRise As Boolean  '�ߴ� �ð� �ִ���
  bSet As Boolean  '���� �ð� �ִ���
End Type

Type ObjPos
  RA1 As Double  '00��
  RA2 As Double  '12��
  RA3 As Double  '24��
  DE1 As Double  '00��
  DE2 As Double  '12��
  DE3 As Double  '24��
End Type

Public Const HorSun As Byte = 4
Public Const HorMoon As Byte = 5

'for planetpos function
Public Const SUN As Byte = 0
Public Const MOON As Byte = 8

Function RSTime(Longi As Double, Lati As Double, jul As Double, ByVal TZ As Double, ByVal T As Byte, ByVal PL As Byte, ByVal PREC As Integer) As String
  Dim P As ObjPos, D As tRSTime, oH As Double, M(2) As Single, jd0 As Double, pName(9) As String, topo As Boolean, P2 As ObjPos
  Dim M2(2) As Single, RSTR As String
    
  Select Case T
   Case HorSun: oH = -0.8333
   Case HorMoon: oH = 0.125
  End Select
    
  '������ ����
  D.Longitude = Longi
  D.Latitude = Lati
  D.ObjAlt = oH
  D.JD = jul
  D.TZone = TZ
  
  '��¥ �� õü ��ġ ���
  jd0 = GetJD0(D.JD)
  Call PlanetPosB(PL, jd0, D.TZone, True, P.RA1, P.DE1, M(0))
  Call PlanetPosB(PL, jd0 + 0.5, D.TZone, True, P.RA2, P.DE2, M(1))
  Call PlanetPosB(PL, jd0 + 1, D.TZone, True, P.RA3, P.DE3, M(2))
  
  GetRiseSetByPos D, P, PREC
  
  '���
  If D.bRise = True Then RSTR = MakeTimeString3(D.RiseTime) Else RSTR = "--:--"
  If D.bSet = True Then RSTR = RSTR & "/" & MakeTimeString3(D.SetTime) Else RSTR = RSTR & "/--:--"

  If PL = MOON Then
    Call PlanetPosB(PL, jd0 + TZ / 24, D.TZone, True, P.RA1, P.DE1, M(1))
    RSTR = RSTR & "(" & Round(M(1) * 100) & "%)"
  End If
  RSTime = RSTR
End Function

'��ġ�� �׳� 00, 12, 24�� �ڷ�� �Է�
Sub GetRiseSetByPos(DataSet As tRSTime, PosData As ObjPos, PREC As Integer)
  Dim jd0 As Double, dt As Double, R As Double, S As Double
  Dim br As Boolean, bs As Boolean
  Dim RA As Double, de As Double, D As Double
  Dim ALT(1440) As Double, az(1440) As Double, T(1440) As Double
  Dim RA1 As Double, RA2 As Double, RA3 As Double, DE1 As Double, DE2 As Double, DE3 As Double
  Dim i As Double, oAlt As Double, La As Double, LO As Double, N As Double, N2 As Double
  
  Dim ut_now As Double, ut0 As Double, temp As Double, t1 As Double
  ut_now = GetJD0(DataSet.JD) - DataSet.TZone / 24
  ut0 = GetJD0(ut_now)
  t1 = (ut0 - 2451545#) / 36525
  temp = Rev(100.46061837 + 36000.770053608 * t1 + 0.000387933 * t1 * t1 - t1 * t1 * t1 / 38710000) + DataSet.Longitude + (ut_now - ut0) * 360.985647366
  temp = Rev(temp)
  
  '������ ����
  LO = DataSet.Longitude
  La = DataSet.Latitude
  oAlt = DataSet.ObjAlt
  
  '��¥ �� õü ��ġ ���
  jd0 = GetJD0(DataSet.JD)
  DE1 = PosData.DE1: DE2 = PosData.DE2: DE3 = PosData.DE3
  RA1 = PosData.RA1: RA2 = PosData.RA2: RA3 = PosData.RA3
  
  'õü�� ������ǥ�� �д����� ���(1440/n)
  N = 1440 / CDbl(PREC): N2 = N / 2
  dt = 1 / N
  For i = 0 To N
    T(i) = jd0 + i * dt
    Inter3Sph RA1, DE1, RA2, DE2, RA3, DE3, (i - N2) / N2, RA, de
    EquToAltAz RA, de, (temp + i * dt * 360.985647366) / 15, La, az(i), ALT(i)
  Next i
  
  '��� ����� �������� ��, �� ã��
  'br: �ߴ� �ð� ����, bs: ���� �ð� ����,bwn: õü�� �Ϸ����� �� ����
  br = False: bs = False
  For i = 0 To (N - 1)
    If ALT(i) <= oAlt And oAlt <= ALT(i + 1) Then
      br = True
      R = Int2(ALT(i), ALT(i + 1), T(i), T(i + 1), oAlt)
    End If
    If ALT(i) >= oAlt And oAlt >= ALT(i + 1) Then
      bs = True
      S = Int2(ALT(i), ALT(i + 1), T(i), T(i + 1), oAlt)
    End If
  Next i
  
  '���
  With DataSet
    .RiseTime = R
    .SetTime = S
    .bRise = br
    .bSet = bs
  End With
End Sub

