Attribute VB_Name = "z_FileDialogue"
Option Explicit

'###############################################################
'�����ο��� VBA ����������Լ� (https://www.oppadu.com)
'�� vbFileSearch �Լ�
'�� ������ ������ Ư�� Ȯ���ڸ� ���� ���ϸ��� ������ġ/��Ȯ�� ��ġ�� �˻��մϴ�.
'�� �μ� ����
'_____________FileName      : �˻��� ���ϸ��Դϴ�.
'_____________LookIn        : ��ȸ�� �����Դϴ�.
'_____________Extension     : Ư�� Ȯ���ڸ� ��ȸ�մϴ�. (��ǥ(,)�� ����)
'_____________ExactMatch    : ��Ȯ����ġ �˻� �����Դϴ�.
'_____________withPath      : True�� ��� ������� ������θ� ����մϴ�.
'�� ���� ��Ÿ ����������Լ�
'_____________SplitFileExt �Լ�
'_____________IsInArray �Լ�
'_____________ListFiles �Լ�
'###############################################################
Function vbFileSearch(FileName As String, _
                        LookIn As String, _
                        Optional ExactMatch As Boolean = True, _
                        Optional Extension As String = "", _
                        Optional withPath As Boolean = True) As String

Dim sFullName As Variant
Dim sExts As Variant: Dim sExt As Variant
Dim vaArr As Variant: Dim vaRtn As Variant
Dim i As Long: Dim j As Long

vbFileSearch = "-1"
If Right(LookIn, 1) <> "\" Then LookIn = LookIn & "\"

vaArr = SplitFileExt(ListFiles(LookIn, False))
vaRtn = IsInArray(FileName, vaArr, ExactMatch, rtnArrayValue, 0)

If TypeName(vaRtn) = "Variant()" Then
    If Extension = "" Then
        If withPath = True Then vbFileSearch = LookIn & vaRtn(0, 0) & vaRtn(0, 1) Else: vbFileSearch = vaRtn(0, 0) & vaRtn(0, 1)
    Else
        sExts = Split(Extension, ",")
        For Each sExt In sExts
            For i = LBound(vaRtn) To UBound(vaRtn)
                If StrComp(CStr(vaRtn(i, 1)), CStr("." & Trim(sExt)), vbTextCompare) = 0 Then
                    If withPath = True Then
                        vbFileSearch = LookIn & vaRtn(i, 0) & vaRtn(i, 1)
                        Exit Function
                    Else
                        vbFileSearch = vaRtn(i, 0) & vaRtn(i, 1)
                        Exit Function
                    End If
                End If
            Next
        Next
        vbFileSearch = "-1"
    End If
Else
    vbFileSearch = vaRtn
End If

End Function

'###############################################################
'�����ο��� VBA ����������Լ� (https://www.oppadu.com)
'�� ListFiles �Լ�
'�� ������ ������ ���ϸ���� �迭�� ��ȯ�մϴ�.
'�� �μ� ����
'_____________sPath     : ���ϸ���� ����� �����Դϴ�.
'_____________withPath  : ������θ� ���� ����� ���θ� �����մϴ�.
'###############################################################

Function ListFiles(sPath As String, Optional withPath As Boolean = False)

'// �� ������ �����մϴ�.
Dim arr As Variant
Dim i As Integer
Dim oFSO As Object: Dim oFolder As Object: Dim oFiles As Object: Dim oFile As Object

Set oFSO = CreateObject("Scripting.FileSystemObject")
Set oFolder = oFSO.GetFolder(sPath)
Set oFiles = oFolder.Files

'// ������ �����մϴ�.
i = 1
If oFiles.Count = 0 Then Exit Function
If Right(sPath, 1) <> "\" Then sPath = sPath & "\"

'// ������ ������ �Ѱ��� ����� �迭 �����մϴ�.
ReDim arr(1 To oFiles.Count)

'// �� ������ ���ư��� Arr �迭�� ��ȯ�մϴ�.
For Each oFile In oFiles
    If withPath = False Then arr(i) = oFile.Name Else: arr(i) = sPath & oFile.Name
    i = i + 1
Next

'// �迭�� ��������� ����մϴ�.
ListFiles = arr

End Function


'###############################################################
'�����ο��� VBA ����������Լ� (https://www.oppadu.com)
'�� SplitFileExt �Լ�
'�� �迭 �Ǵ� String���� �޾ƿ� Ȯ���ڰ� ���Ե� ���ϰ�θ� �����̸��� Ȯ���ڷ� �и��մϴ�. (2���� �迭, 0 - ���ϸ�, 1 - Ȯ����)
'�� �μ� ����
'_____________Files     : Ȯ���ڸ� �и��� �迭 �Ǵ� ���ϸ��Դϴ�.
'###############################################################
Function SplitFileExt(Files As Variant) As Variant

Dim i As Long
Dim sFile As String
Dim vaArr As Variant

On Error GoTo ErrHandler:

'// ���� Ÿ���� Ȯ���մϴ�. �迭 �Ǵ� ��Ÿ ���ڿ��� ���
If TypeName(Files) = "Variant()" Then
    '// �Է� ������ �迭�� ���
    ReDim vaArr(LBound(Files) To UBound(Files), 0 To 1)
    For i = LBound(Files) To UBound(Files)
        sFile = Files(i)
        vaArr(i, 0) = Left(sFile, InStrRev(sFile, ".") - 1)
        vaArr(i, 1) = Right(sFile, Len(sFile) - InStrRev(sFile, ".") + 1)
    Next
Else
    '// �Է� ������ ��Ÿ ���ڿ��� ���
    ReDim vaArr(0, 1)
    vaArr(0, 0) = Left(Files, InStrRev(Files, ".") - 1)
    vaArr(0, 1) = Right(Files, Len(Files) - InStrRev(Files, ".") + 1)
End If

SplitFileExt = vaArr


Exit Function

ErrHandler:
MsgBox "�ùٸ� ���ϰ�� �Ǵ� ���ϸ��� �Է��ϼ���." & sFile
End

End Function
