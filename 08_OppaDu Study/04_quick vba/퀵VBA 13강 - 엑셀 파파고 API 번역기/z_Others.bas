Attribute VB_Name = "z_Others"
Option Explicit

'---------------------------------------------------
' ���ó�¥ Ȯ�� �� API ��뷮 �ʱ�ȭ
'---------------------------------------------------
Sub Auto_Open()
'���������� �α��� �� ��¥�� ���� ��¥�� �ٸ� ��� API ��뷮�� �ʱ�ȭ�մϴ�.
If Sheet2.Range("U8").Value <> Date Then
    Sheet2.Range("U8").Value = Date
    Sheet2.Range("U6").Value = 0
End If
End Sub

'---------------------------------------------------
' Ŭ������ ����/�ٿ��ֱ� �� ���� �ʱ�ȭ ��ɹ�
'---------------------------------------------------
Sub RefreshAll()
'Ŭ�����忡 ����� ���� [����] ��Ʈ �ȿ� �ؽ�Ʈ ���·� �ٿ��ֱ� �� ��
'������ ������Ʈ �մϴ�.
Dim Rng As Range
Application.ScreenUpdating = False

With Sheet1
    Set Rng = .UsedRange
    If Rng.Rows.Count > 1 Then .Range("2:" & Rng.Rows.Count).EntireRow.Delete
    .Activate
    .Range("A2").Select
    If Application.LanguageSettings.LanguageID(msoLanguageIDUI) = 1042 Then
        .PasteSpecial Format:="�����ڵ� �ؽ�Ʈ"
    Else
        .PasteSpecial Format:="Unicode Text"
    End If
End With
Sheet2.Activate
Application.ScreenUpdating = True
ThisWorkbook.RefreshAll

End Sub

'---------------------------------------------------
' ������ȯ ���� ��ɹ�
'---------------------------------------------------
Sub vocStop()
' �������� ������ȯ�� �ߴ��մϴ�.
VocSpeak blnStop:=True

End Sub
Sub vocSpeak_To()
' ������ �ؽ�Ʈ�� �������� ��ȯ�մϴ�.
Dim x As Long: Dim i As Long
Dim s As String
With Sheet2
    x = .Range("P14").End(xlDown).Row
    If x = .Rows.Count Then x = 14
    For i = 14 To x
        s = s & .Range("P" & i).Value
    Next
    VocSpeak s, Application.WorksheetFunction.VLookup(Sheet2.Range("P9").Value, Sheet3.Range("A:C"), 3, 0)
End With
End Sub
Sub vocSpeak_From()
' �����Ǳ� �� �ؽ�Ʈ�� �������� ��ȯ�մϴ�.
Dim x As Long: Dim i As Long
Dim s As String
With Sheet2
    x = .Range("F14").End(xlDown).Row
    If x = .Rows.Count Then x = 14
    For i = 14 To x
        s = s & .Range("F" & i).Value
    Next
    VocSpeak s, Application.WorksheetFunction.VLookup(Sheet2.Range("F9").Value, Sheet3.Range("A:C"), 3, 0)
End With
End Sub
