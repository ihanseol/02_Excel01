VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmShtSelect 
   Caption         =   "��Ʈ�� �����ϼ���."
   ClientHeight    =   3180
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   2205
   OleObjectBlob   =   "frmShtSelect.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmShtSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnSubmit_Click()

Dim WB As Workbook
Dim WS As Worksheet: Dim NewWS As Worksheet
Dim rng As Range
Dim i As Long: i = 0: Dim j As Long: j = 2
Dim endCol As Long: Dim endRow As Long
Dim strWS As String

Set WB = ThisWorkbook

'// ����Ʈ�ڽ� ���ÿ��� Ȯ��
If isListBoxSelected(Me.lstSheet) = False Then MsgBox "��Ʈ�� �����ϼ���": Exit Sub


'// ��Ʈ���� ��Ʈ ���翩�θ� Ȯ���մϴ�.
For Each WS In WB.Worksheets
    If WS.Name = "��Ʈ����" Then MsgBox "��Ʈ���� ��Ʈ�� �����մϴ�.": Exit Sub
Next

Set NewWS = WB.Worksheets.Add(before:=WB.Worksheets(1))
NewWS.Name = "��Ʈ����"
With NewWS
    .Cells(1, 1) = "����"
    .Cells(1, 2) = "��¥"
    .Cells(1, 3) = "�̸�"
    .Cells(1, 4) = "��ٽð�"
    .Cells(1, 5) = "��ٽð�"
End With

'// ����Ʈ�ڽ����� ���õ� ��Ʈ�� �޾ƿ���
For i = 0 To Me.lstSheet.ListCount - 1
    If Me.lstSheet.Selected(i) = True Then
        strWS = Me.lstSheet.List(i)
        Set WS = WB.Worksheets(strWS)
        With WS
            endCol = .Cells(1, .Columns.Count).End(xlToLeft).Column
            endRow = .Cells(.Rows.Count, 1).End(xlUp).Row
            Set rng = .Range(.Cells(2, 1), .Cells(endRow, endCol))
            rng.Copy NewWS.Cells(j, 1)
            j = j + rng.Rows.Count
        End With
    End If
Next

'// ��Ʈ���� �Ϸ� �ȳ�
MsgBox "��Ʈ������ �Ϸ�Ǿ����ϴ�."
Unload Me

End Sub

Private Sub UserForm_Initialize()

Dim WB As Workbook
Dim WS As Worksheet

Set WB = ThisWorkbook

For Each WS In WB.Worksheets
    Me.lstSheet.AddItem WS.Name
Next

End Sub

Function isListBoxSelected(ListBox As MSForms.ListBox) As Boolean

Dim i As Long

For i = 0 To ListBox.ListCount - 1
    If ListBox.Selected(i) Then isListBoxSelected = True: Exit Function
Next

isListBoxSelected = False

End Function
