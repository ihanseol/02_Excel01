VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmShtSelection 
   Caption         =   "��Ʈ�� �����ϼ���."
   ClientHeight    =   3216
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   2910
   OleObjectBlob   =   "frmShtSelection.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmShtSelection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnSubmit_Click()

Dim WB As Workbook
Dim WS As Worksheet: Dim newWS As Worksheet
Dim Rng As Range
Dim i As Long: Dim j As Long: j = 2
Dim endCol As Long: Dim endRow As Long

Set WB = ThisWorkbook

'// isListBoxSelected �Լ��� ���� �ڼ��� ������ �Ʒ� ��ũ�� �����ϼ���.
'// https://oppadu.com/����-����/vba-����Ʈ�ڽ�-��-���ÿ���/
If isListBoxSelected(Me.lstSheet) = False Then MsgBox "��Ʈ�� �����ϼ���.": Exit Sub

For Each WS In WB.Worksheets
    If WS.Name = "��Ʈ����" Then MsgBox "��Ʈ���� ��Ʈ�� �����մϴ�. ��Ʈ���� ��Ʈ�� �̸��� �����Ͻðų� �����Ͻ� �� �ٽ� �������ּ���.": Exit Sub
Next

Set newWS = WB.Worksheets.Add(before:=WB.Worksheets(1))
newWS.Name = "��Ʈ����"
With newWS
    .Cells(1, 1) = "����"
    .Cells(1, 2) = "��¥"
    .Cells(1, 3) = "�̸�"
    .Cells(1, 4) = "��ٽð�"
    .Cells(1, 5) = "��ٽð�"
End With

For i = 0 To Me.lstSheet.ListCount - 1
    If Me.lstSheet.Selected(i) = True Then
        Set WS = WB.Worksheets(Me.lstSheet.List(i))
        With WS
            '// ��Ʈ�� ������ ��/�� �޾ƿ��⿡ ���� �ڼ��� ������ �Ʒ� ��ũ�� �����ϼ���.
            '// https://oppadu.com/����-����/����-vba-������-��-ã��-������-��-ã��/
            endRow = .Cells(.Rows.Count, 1).End(xlUp).Row
            endCol = .Cells(1, .Columns.Count).End(xlToLeft).Column
            Set Rng = .Range(.Cells(2, 1), .Cells(endRow, endCol))
            Rng.Copy newWS.Cells(j, 1)
            j = j + Rng.Rows.Count
        End With
    End If
Next

MsgBox "��Ʈ ������ �Ϸ�Ǿ����ϴ�."
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
