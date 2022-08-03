Attribute VB_Name = "z_Array"
Option Explicit

Public Enum xlArrayReturnType
    rtnSequence = 1
    rtnValue = 2
    rtnArrayValue = 3
End Enum

'###############################################################
'�����ο��� VBA ����������Լ� (https://www.oppadu.com)
'�� IsInArray �Լ�
'�� �迭�ȿ� ������ ���� ���翩�θ� Ȯ���մϴ�. 4���� ������ ��ȯ �����մϴ�. (����, ��, �迭)
'�� �μ� ����
'_____________FindValue     : �迭���� ã�� ���Դϴ�.
'_____________vaArray       : ���� �˻��� �迭�Դϴ�.
'_____________ExactMatch    : ��Ȯ����ġ/������ġ �˻������Դϴ�.
'_____________returnType    : ��ȯ������ �����մϴ�. (Public Enum)
'_____________Dimension     : ���� �˻��� �迭�� �����Դϴ�.
'_____________vbCompare     : ��ҹ��� ���� �����Դϴ�. (vbTextCompare : ��ҹ��ڱ���X, vbBinaryCompare : ��ҹ��ڱ���)
'�� ���� ��Ÿ ����������Լ�
'_____________ArrayDimension �Լ�
'�� �׿� �������
'_____________xlArrayReturnType PublicEnum�� ���Ǿ����ϴ�.
'###############################################################

Function IsInArray(FindValue As Variant, _
                    vaArray As Variant, _
                    Optional ExactMatch As Boolean = True, _
                    Optional returnType As xlArrayReturnType = rtnValue, _
                    Optional Dimension As Integer = 0, _
                    Optional vbCompare As VbCompareMethod = vbTextCompare) As Variant

Dim i As Long: Dim j As Long
Dim dicArr As Object: Dim dicKey As Variant: Dim dicKeys As Variant
Dim rtnArr As Variant
Dim ArrDim As Integer

ArrDim = ArrayDimension(vaArray)
Set dicArr = CreateObject("Scripting.Dictionary")

'// IsInArray �ʱⰪ�� �����մϴ�.
IsInArray = -1

'// ��Ȯ����ġ�� ���
If ExactMatch = True Then
        '// ���� ������ ReturnType�� ���� ��ȯ�մϴ�. (��Ȯ�� ��ġ�̹Ƿ� ����� ������ 1���� ��ȯ)
        If ArrDim = 1 Then
        '// �迭�� 1������ ���
            For i = LBound(vaArray) To UBound(vaArray)
                If StrComp(FindValue, vaArray(i), vbCompare) = 0 Then
                    If returnType = rtnValue Or returnType = rtnArrayValue Then IsInArray = vaArray(i): Exit For
                    If returnType = rtnSequence Then IsInArray = i: Exit For
                End If
            Next i
        Else
        '// �迭�� 2������ ���
           For i = LBound(vaArray) To UBound(vaArray)
                If StrComp(FindValue, vaArray(i, Dimension), vbCompare) = 0 Then
                    If returnType = rtnValue Then IsInArray = vaArray(i, Dimension): Exit For
                    If returnType = rtnSequence Then IsInArray = i: Exit For
                    If returnType = rtnArrayValue Then dicArr.Add i, i: Exit For
                    Exit For
                End If
            Next i
            
            '// Dictionary�� ���� �ϳ��� �����
            If dicArr.Count > 0 Then
                '�迭 Redim
                ReDim rtnArr(0, 0 To ArrDim - 1)
                    '// Dictionary ���� �޾ƿ� �� ���� �迭�� �ű�ϴ�
                    For Each dicKey In dicArr.keys
                        For j = LBound(vaArray, 2) To UBound(vaArray, 2)
                            rtnArr(0, j) = vaArray(dicKey, j)
                        Next
                    Next
                '// �迭�� ����� ����մϴ�
                IsInArray = rtnArr
            End If
        End If
   
'// ������ġ�� ���
Else
    If ArrDim = 1 Then
        '// �迭�� 1������ ���
        For i = LBound(vaArray) To UBound(vaArray)
        '// ���� ������ ReturnType�� ���� ����� ��ȯ�մϴ�
            If InStr(1, vaArray(i), FindValue) > 0 Then
                If returnType = rtnValue Then IsInArray = vaArray(i): Exit For
                If returnType = rtnSequence Then IsInArray = i: Exit For
                If returnType = rtnArrayValue Then dicArr.Add i, i
            End If
        Next i
        
        '// ������ġ ���� �ϳ��� �����
        If dicArr.Count > 0 Then
            '�迭 Redim
            ReDim rtnArr(0 To dicArr.Count - 1)
            i = 0
                '// Dictionary ���� �޾ƿ� �� ���� �迭�� �ű�ϴ�
                For Each dicKey In dicArr.keys
                        rtnArr(i) = vaArray(dicKey)
                    i = i + 1
                Next
            '// �迭�� ����� ����մϴ�
            IsInArray = rtnArr
        End If
    Else
        '// �迭�� 2������ ���
        For i = LBound(vaArray) To UBound(vaArray)
        '// ���� ������ ReturnType�� ���� ����� ��ȯ�մϴ�
            If InStr(1, vaArray(i, Dimension), FindValue) > 0 Then
                If returnType = rtnValue Then IsInArray = vaArray(i, Dimension): Exit For
                If returnType = rtnSequence Then IsInArray = i: Exit For
                If returnType = rtnArrayValue Then dicArr.Add i, i
            End If
        Next i
        
        '// ������ġ ���� �ϳ��� �����
        If dicArr.Count > 0 Then
            '�迭 Redim
            ReDim rtnArr(0 To dicArr.Count - 1, 0 To ArrDim - 1)
            i = 0
                '// Dictionary ���� �޾ƿ� �� ���� �迭�� �ű�ϴ�
                For Each dicKey In dicArr.keys
                    For j = LBound(vaArray, 2) To UBound(vaArray, 2)
                        rtnArr(i, j) = vaArray(dicKey, j)
                    Next
                    i = i + 1
                Next
            '// �迭�� ����� ����մϴ�
            IsInArray = rtnArr
        End If
        
        End If
End If
  
End Function


'###############################################################
'�����ο��� VBA ����������Լ� (https://www.oppadu.com)
'�� ArrayDimension �Լ�
'�� �迭�� �������� ��ȯ�մϴ�.
'�� �μ� ����
'_____________vaArray     : ������ ������ �迭�� �Է��մϴ�.
'###############################################################
Function ArrayDimension(vaArray As Variant) As Integer

Dim i As Integer: Dim x As Integer

On Error Resume Next

Do
    i = i + 1
    x = UBound(vaArray, i)
Loop Until Err.Number <> 0

Err.Clear

ArrayDimension = i - 1

End Function

