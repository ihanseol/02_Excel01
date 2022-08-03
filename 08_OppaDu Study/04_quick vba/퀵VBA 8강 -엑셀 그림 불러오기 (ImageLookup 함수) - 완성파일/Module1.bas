Attribute VB_Name = "Module1"
Option Explicit

'###############################################################
'�����ο��� VBA ����� �����Լ� (https://www.oppadu.com)
'�� ImageLookup �Լ� (��ũ��Ʈ �Լ�)
'�� ������ο� �׸������� �����̸��� �����Ͽ� '��Ȯ����ġ' �Ǵ� '������ġ'�� �ش缿�� �̹����� ����մϴ�. (�޸�Ȱ��)
'�� �μ� ����
'_____________ImageName     : �˻��� �̹��������� �̸� (Ȯ���� ���� X)
'_____________FolderPath    : �̹����� �˻��� �������
'_____________ExactMatch    : �׸����� �̸� ��Ȯ�� ��ġ �˻�����
'_____________NAValue       : �׸����� ���� �� ����� ���� �޼���
'_____________ShowImageName : �׸����� ��� ��¿���
'_____________ShowBorder    : �׸������� �׵θ� ��¿���
'_____________Extension     : Ư�� �׸����� ���ĸ� �˻��Ұ��, ���ϴ� Ȯ���� ����
'�� ���� ��Ÿ ����������Լ�
'_____________cvRng �Լ�
'_____________vbFileSearch �Լ�
'�� �׿� �������
'###############################################################

Function ImageLookup(ImageName, _
                        Optional FolderPath = "", _
                        Optional ExactMatch As Boolean = True, _
                        Optional NAvalue As String = "-", _
                        Optional ShowImageName As Boolean = False, _
                        Optional ShowBorder As Boolean = False, _
                        Optional Extension = "png, jpeg, jpg, gif")
                        
                        
Dim rng As Range
Dim myComment As Comment
Dim sName As String: Dim FullPath As String

'// �Լ��� �Էµ� ��
Set rng = Application.Caller
sName = CStr(cvRng(ImageName))

'// �Լ��� �Էµ� ���� �޸� ����
rng.ClearComments

'// ������ �����̸��� ��ĭ�� ��� �����޼��� ��ȯ
If sName = "" Then ImageLookup = NAvalue: Exit Function

On Error GoTo EmptyFolder:
'// ������� ��ĭ�� ��� �ش� ��ũ��Ʈ ������� ��ȯ
If FolderPath = "" Then FolderPath = rng.Parent.Parent.Path Else: FolderPath = cvRng(FolderPath)
On Error GoTo 0

'// ������ξȿ� ���� �˻�
FullPath = vbFileSearch(sName, CStr(FolderPath), ExactMatch, CStr(Extension))

'// �����ȿ� ���� ����� �Լ��� �Էµ� ���� �޸���� �� �̹��� ���
If FullPath <> "-1" Then
    Set myComment = rng.AddComment
    
    With myComment
        .Visible = True

        If sName = "" Or ShowImageName = False Then
            .Text Text:=" "
        ElseIf ShowImageName = True Then
            .Text Text:=FullPath
        End If
    
        With .Shape
                .Left = rng.Left + 3
                .Top = rng.Top + 3
    
    
                .Width = rng.MergeArea.Width - 6
                .Height = rng.MergeArea.Height - 6
                
                .Fill.UserPicture FullPath
                .Line.ForeColor.RGB = rng.Interior.Color
                
                If ShowBorder = True Then .Line.Visible = msoTrue Else: .Line.Visible = msoFalse
        End With
    End With
    ImageLookup = ""
Else
'// �����ȿ� ���� ���� ��� �����޼��� ���
    ImageLookup = NAvalue
End If

Exit Function

EmptyFolder:
ImageLookup = NAvalue: Exit Function

End Function

