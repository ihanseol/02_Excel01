VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmWBSelect 
   Caption         =   "������ ������ �ּ���."
   ClientHeight    =   2220
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   7170
   OleObjectBlob   =   "frmWBSelect.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmWBSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnMerge_Click()

Dim WB As Workbook
Dim WS As Worksheet: Dim toWS As Worksheet
Dim rng As Range
Dim i As Long: i = 0: Dim j As Long
Dim endCol As Long: Dim endRow As Long
Dim strWS As String

'// ��ũ��������Ʈ �ߴ�
Application.ScreenUpdating = False
Application.DisplayAlerts = False

'// ��������
If Me.lstWB.ListCount = 0 Then
    MsgBox "������ ������ �����ϼ���."
    Exit Sub
End If

'// ���Ϻ���
Set toWS = ActiveSheet
j = toWS.Cells(toWS.Rows.Count, 1).End(xlUp).Row

For i = 0 To Me.lstWB.ListCount - 1
    Set WB = Application.Workbooks.Open(Me.lstWB.List(i))
    For Each WS In WB.Worksheets
        If WS.Name Like Me.txtFilter.Value & "*" Then
                With WS
                    endCol = .Cells(1, .Columns.Count).End(xlToLeft).Column
                    endRow = .Cells(.Rows.Count, 1).End(xlUp).Row
                    Set rng = .Range(.Cells(2, 1), .Cells(endRow, endCol))
                    rng.Copy toWS.Cells(j, 1)
                    j = j + rng.Rows.Count
                End With
        End If
    Next
    WB.Close
Next

'// �ȳ��޼���
MsgBox "���� ������ �Ϸ� �Ǿ����ϴ�."
Unload Me

'//��ũ�� ������Ʈ Ȱ��ȭ
Application.ScreenUpdating = True
Application.DisplayAlerts = True

End Sub

Private Sub btnSelect_Click()

Dim strFilePath As String
Dim varFilePaths As Variant: Dim varFilePath As Variant

strFilePath = Multiple_FileDialog

varFilePaths = Split(strFilePath, ", ")

Me.lstWB.Clear

For Each varFilePath In varFilePaths
    Me.lstWB.AddItem varFilePath
Next

End Sub
