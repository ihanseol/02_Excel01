VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} RefChkBoxHwaSoo 
   Caption         =   "Check Anything"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "RefChkBoxHwaSoo.frx":0000
   StartUpPosition =   1  '������ ���
End
Attribute VB_Name = "RefChkBoxHwaSoo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandButton1_Click()

    Dim DBSheet As Worksheet
    Dim w() As String
    Dim Singlew As Variant
    Dim i As Integer
    Dim j As Integer
    Dim r As Integer
    Dim a As Integer
    Dim InputYear As Integer
    Dim v() As String
    Dim CalendarData() As Variant
    Dim dataMonth As Integer
    Dim SingleRange As Range
    Dim dDate As Date
    Dim eMon, eDay As Integer '���� ��,�� ����
    Dim strDate As String
    
    '�޷� �ʱ�ȭ
    Call ClearCalendar

    Set DBSheet = Sheets("DB")
    InputYear = year(Date)

    
    'üũ�ڽ� �ε����� �迭�� ����
    For i = 0 To Me.Controls.Count - 1
        If Me.Controls(i) Then
            ReDim Preserve w(j)
            w(j) = i
            j = j + 1
        End If
    Next
    
    'üũ�ڽ� ���� �� �ϸ� ��ũ�� ����
    If j <= 1 Then GoTo j

    'üũ�� ������ �� ��¥�� ���� Ȯ���Ͽ� j�� ����
    j = 0

    For Each Singlew In w
        If DBSheet.Cells(1, CInt(Singlew)).Value = "����" Then
            j = Singlew
        End If
    Next
    
    '��¥���� �������� ������ ��ũ�� ����
    If j = 0 Then
        MsgBox "��¥�� �ִ� ���� �����ϼ���", vbCritical
        GoTo j
    End If
    
    '������ ���� ��ȯ�ϸ� �޷¿� ����� �����͸� �迭�� ����
    r = DBSheet.Range("A1").CurrentRegion.Rows.Count
    ReDim v(UBound(w) - 1)
    ReDim CalendarData(r - 2, 1)
    For i = 2 To r
        For Each Singlew In w
            If Singlew <> j Then
                v(a) = DBSheet.Cells(i, CInt(Singlew)).Value
                a = a + 1
            End If
        Next
        a = 0
        
        '����� �����Ϳ� ��¥�� �и��Ͽ� �迭�� ����
        CalendarData(i - 2, 0) = Join(v, ", ")
        
        If DBSheet.Cells(i, "D").Value = "����" Then
            dDate = dependOnMonthType(DBSheet.Cells(i, CInt(j)).Value)
            eMon = Month(dDate)
            eDay = Day(dDate)
            
            InputYear = IIf(eMon = 12, year(Date) - 1, year(Date))
            strDate = Lun2Sol(InputYear, eMon, eDay, isYoonDal2(dDate))
            
            CalendarData(i - 2, 1) = getYMD(strDate)
        Else
            dDate = dependOnMonthType(DBSheet.Cells(i, CInt(j)).Value)
            CalendarData(i - 2, 1) = dDate
        End If
    Next
    
    '������ �迭�� ��ȯ�Ͽ� �޷¿� ������ ���
    For i = 0 To r - 2
        dDate = dependOnMonthType(CalendarData(i, 1))
        dataMonth = Month(dDate) + 2
        
        'dataMonth = Month(CalendarData(i, 1)) + 2
        
        For Each SingleRange In Sheets(dataMonth).UsedRange
            'If SingleRange = CalendarData(i, 1) Then
             If SingleRange = dDate Then
                With SingleRange.Offset(1)
                    If .Value <> "" Then
                        .Value = .Value & vbLf & CalendarData(i, 0)
                    Else
                        .Value = CalendarData(i, 0)
                    End If
                    .HorizontalAlignment = xlLeft
                    .Font.Color = vbBlack
                End With
            End If
        Next
        
    Next

j:
    Unload Me
    
End Sub


Function isYoonDal(d As Date) As Boolean
    Dim arr As Variant
    Dim yy, mm As Integer
    Dim thisYear As Integer
    Dim i, j As Integer
    
    thisYear = year(Date)
    yy = year(d)
    mm = Month(d)
    
    arr = [{2023, 2 ; 2025,6 ; 2028,5 ; 2031,3 ; 2033,11 ; 2036,6 ; 2039,5 ; 2042,2 ; 2044,7 ; 2047,5 ; 2050,3 ; 2052,8 ; 2055,6 ; 2058,4 ; 2061,3}]
    
    For i = 1 To UBound(arr, 1)
        If yy < arr(i, 1) Then
            isYoonDal = False
            Exit Function
        End If
    
        If yy = arr(i, 1) Then
              If mm = arr(i, 2) Then
                isYoonDal = True
                Exit Function
              Else
                isYoonDal = False
                Exit Function
              End If
        End If
    Next i
End Function


Function isYoonDal2(d As Date) As Boolean
    Dim arr As Variant
    Dim yy, mm As Integer
    Dim thisYear As Integer
    Dim i, j As Integer
    Dim dict As New Scripting.Dictionary
    
    thisYear = year(Date)
    yy = year(d)
    mm = Month(d)
    
    arr = [{ 2023, 2 ; 2025,6 ; 2028,5 ; 2031,3 ; 2033,11 ; 2036,6 ; 2039,5 ; 2042,2 ; 2044,7 ; 2047,5 ; 2050,3 ; 2052,8 ; 2055,6 ; 2058,4 ; 2061,3}]
    
    For i = 1 To UBound(arr, 1)
        dict.Add arr(i, 1), arr(i, 2)
    Next i
    
    If dict.Exists(yy) Then
        If mm = dict(yy) Then
          isYoonDal2 = True
        Else
          isYoonDal2 = False
        End If
    Else
        isYoonDal2 = False
    End If
   
End Function





Function getYMD(strDate As String) As Date

    Dim Match As Object
    Dim i As Integer
    Dim InputYear As Integer
    
    InputYear = year(Date)
    
    With CreateObject("VBscript.regexp")    '���Խ� ����
    
        .Global = True  '��� ���� �ľ�
        .Pattern = "\d+"    '���ڸ�
        
        If .test(strDate) = True Then
            Set Match = .Execute(strDate)
             getYMD = DateSerial(CInt(Match(0)), CInt(Match(1)), CInt(Match(2)))
        Else
            getYMD = Date
        End If
        
    End With


End Function


Function dependOnMonthType(ByVal d As Variant) As Date
    
    'Debug.Print TypeName(d)
    
    If TypeName(d) = "String" Then
        dependOnMonthType = makeDateFromString(d)
    Else
        dependOnMonthType = d
    End If

End Function

Function makeDateFromString(tmp As Variant) As Date
 
    Dim Match As Object
    Dim i As Integer
    Dim InputYear As Integer
    
    InputYear = year(Date)
    
    With CreateObject("VBscript.regexp")    '���Խ� ����
    
        .Global = True  '��� ���� �ľ�
        .Pattern = "\d+"    '���ڸ�
        
        If .test(tmp) = True Then
            Set Match = .Execute(tmp)
             makeDateFromString = DateSerial(InputYear, CInt(Match(0)), CInt(Match(1)))
        Else
            makeDateFromString = Date
        End If
        
    End With
    
End Function
 



Private Sub UserForm_Initialize()

    Dim i As Integer
    Dim OptionList
    Dim chkBox As MSForms.CheckBox
    Dim btn As CommandButton
    Dim DBSheet As Worksheet
    
    '�޷¿� ����� ���ϴ� ������ üũ�ڽ��� ��Ÿ����
    Set DBSheet = Sheets("DB")
    OptionList = DBSheet.Range("A1").CurrentRegion.Rows(1)
    
    For i = 1 To UBound(OptionList, 2)
    
        Set chkBox = Me.Controls.Add("Forms.CheckBox.1", "CheckBox_" & i)
        
        With chkBox
            .Caption = OptionList(1, i)
            .Left = 5
            .Top = 5 + ((i - 1) * 20)
            
            Me.Width = .Width
            Me.Height = .Height * (i + 2)
        End With
    
    Next i
    
    'Ȯ�ι�ư ����
    Set btn = Me.CommandButton1
    
    With btn
        .Caption = "Ȯ��"
        .Top = Me.Height - .Height + (0.5 * chkBox.Height) + 10
        .Left = (Me.Width * 0.5) - (.Width * 0.5)
        Me.Height = Me.Height + .Height + (0.5 * chkBox.Height) + 20
    End With
    
End Sub
