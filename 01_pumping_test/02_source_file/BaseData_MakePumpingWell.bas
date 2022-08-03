Attribute VB_Name = "BaseData_MakePumpingWell"

Option Explicit


'��Ʈ�� �����Ҷ����� ��ü ��������Ÿ�� �ǵ��� �ʰ�, �켱���� ��Ʈ������ �����°��� �⺻���� ������ �ִ�.

Public Sub CopyOneSheet()

    Dim n_sheets As Integer

    n_sheets = sheets_count()
    
    '2020/5/30 ��������Ʈ�� ��ϻ������ִ� �κ� �߰�
    InsertOneRow (n_sheets)
    
    
    If (n_sheets = 1) Then
        Sheets("1").Select
        Sheets("1").Copy Before:=Sheets("Q1")
        
        ActiveSheet.Shapes.Range(Array("CommandButton1")).Select
        Selection.Delete
        ActiveSheet.Shapes.Range(Array("CommandButton2")).Select
        Selection.Delete
        ActiveSheet.Shapes.Range(Array("CommandButton3")).Select
        Selection.Delete
    Else
        Sheets("2").Select
        Sheets("2").Copy Before:=Sheets("Q1")
    End If
    
    ActiveSheet.Name = CStr(n_sheets + 1)
    Range("b2").value = "W-" & (n_sheets + 1)
    Range("e15").value = CStr(n_sheets + 1)
      
    
    If n_sheets = 1 Then
        Call ChangeCellData(n_sheets + 1, 1)
    Else
        Call ChangeCellData(n_sheets + 1, 2)
    End If
    
    Sheets("Well").Select
End Sub

Private Sub InsertOneRow(ByVal n_sheets As Integer)

    n_sheets = n_sheets + 4
    Rows(CStr(n_sheets) & ":" & CStr(n_sheets)).Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    
    Rows(CStr(n_sheets - 1) & ":" & CStr(n_sheets - 1)).Select
    Selection.Copy
    Rows(CStr(n_sheets) & ":" & CStr(n_sheets)).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False

End Sub

Private Sub ChangeCellData(ByVal nsheet As Integer, ByVal nselect As Integer)
'
' change sheet data direct to well sheet data value
'
'https://stackoverflow.com/questions/18744537/vba-setting-the-formula-for-a-cell

    Range("C2, C3, C4, C5, C6, C7, C8, C15, C16, C17, C18, C19, E17, F21").Select
    Range("F21").Activate
    
    nsheet = nsheet + 3
    Selection.Replace What:=CStr(nselect + 3), Replacement:=CStr(nsheet), LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
 
    
    Range("E21").Select
    'Range("E21").Formula = "=Well!" & Cells(nsheet, 9).Address
    Range("E21").Formula = "=Well!" & Cells(nsheet, "I").Address
    
End Sub



'������ ��Ʈ�� ��ȸ�ϸ鼭, ���� �������� ���߾��ش�.
'
Public Sub JojungSheetData()

    Dim n_sheets As Integer
    Dim i As Integer

    n_sheets = sheets_count()
    
    For i = 1 To n_sheets
    
        Sheets(CStr(i)).Activate
        Call ChangeCellData2(i)
    Next i
    
    
End Sub

Private Sub ChangeCellData2(ByVal nsheet As Integer)

    Dim nselect As String

    Range("C2, C3, C4, C5, C6, C7, C8, C15, C16, C17, C18, C19, E17, F21").Select
    Range("F21").Activate

    nsheet = nsheet + 3
    nselect = Mid(Range("c2").Formula, 8)
    
    'Debug.Print Mid(Range("c2").Formula, 8) & ":" & nselect

    Selection.Replace What:=nselect, Replacement:=CStr(nsheet), LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False


    Range("E21").Select
    Range("E21").Formula = "=Well!" & Cells(nsheet, "I").Address
    
End Sub














