Attribute VB_Name = "sub_FileSelection"
Option Explicit

Public Function Multiple_FileDialog(Optional Title As String = "������ �����ϼ���", Optional FilterName As String = "��������", _
Optional FilterExt As String = "*.xls; *.xlsx; *.xlsm", Optional InitialFolder As String = "", _
Optional InitialView As MsoFileDialogView = msoFileDialogViewList, Optional MultiSelection As Boolean = True) As String

Dim FDG As FileDialog
Dim Selected As Integer: Dim i As Integer
Dim ReturnStr As String

Set FDG = Application.FileDialog(msoFileDialogFilePicker)

With FDG
    .Title = Title
    .Filters.Add FilterName, FilterExt
    .InitialView = InitialView
    .InitialFileName = InitialFolder
    .AllowMultiSelect = MultiSelection
    Selected = .Show

    If Selected = -1 Then
        For i = 1 To FDG.SelectedItems.Count - 1
            ReturnStr = ReturnStr & FDG.SelectedItems(i) & ", "
        Next i
        ReturnStr = ReturnStr & FDG.SelectedItems(.SelectedItems.Count)
        
        Multiple_FileDialog = ReturnStr
    ElseIf Selected = 0 Then
        MsgBox "���õ� ������ �����Ƿ� ���α׷��� �����մϴ�."
        End
    End If
    
End With

End Function

Sub OpenFiles()

Dim SelectionStr As String
Dim Vars As Variant: Dim Var As Variant

SelectionStr = Multiple_FileDialog

Vars = Split(SelectionStr, ", ")

For Each Var In Vars
    Application.Workbooks.Open Var
Next

MsgBox "���õ� ���� ������ ��� �����Ͽ����ϴ�."

End Sub

