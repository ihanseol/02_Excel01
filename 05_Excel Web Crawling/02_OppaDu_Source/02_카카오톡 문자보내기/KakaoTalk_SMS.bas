Attribute VB_Name = "KakaoTalk_SMS"
Option Explicit

#If VBA7 Then
' https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-findwindowa
Private Declare PtrSafe Function FindWindow Lib "user32.dll" Alias "FindWindowA" _
                                                        (ByVal lpClassName As String, _
                                                        ByVal lpWindowName As String) As Long
' https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-findwindowexa
Private Declare PtrSafe Function FindWindowEx Lib "user32.dll" Alias "FindWindowExA" _
                                                            (ByVal hwndParent As Long, _
                                                            ByVal hwndChildAfter As Long, _
                                                            ByVal lpszClass As String, _
                                                            ByVal lpszWindow As String) As Long
' https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-sendmessagea
Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" _
                                                            (ByVal hwnd As Long, _
                                                            ByVal wMsg As Long, _
                                                            ByVal wParam As Long, _
                                                            ByRef lParam As Any) As Long
'https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-postmessagea
Private Declare PtrSafe Function PostMessage Lib "user32" Alias "PostMessageA" _
                                                            (ByVal hwnd As Long, _
                                                            ByVal wMsg As Long, _
                                                            ByVal wParam As Long, _
                                                            ByRef lParam As Any) As Long
'https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-keybd_event
Private Declare PtrSafe Sub keybd_event Lib "user32.dll" _
                                                            (ByVal bVk As Byte, _
                                                            ByVal bScan As Byte, _
                                                            ByVal dwFlags As Long, _
                                                            ByVal dwExtraInfo As Long)
'https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-getkeystate
Private Declare PtrSafe Function GetKeyState Lib "user32" ( _
                                                            ByVal nVirtKey As Long) As Long
'https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-getwindow
Public Declare PtrSafe Function GetWindow _
                                                        Lib "user32" _
                                                        (ByVal hwnd As Long, _
                                                        ByVal wCmd As Long) As Long
'https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-getclassname
Private Declare PtrSafe Function GetClassName Lib "user32" Alias "GetClassNameA" _
                                                            (ByVal hwnd As Long, _
                                                            ByVal lpClassName As String, _
                                                            ByVal nMaxCount As Long) As Long
'https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-getwindowtexta
Private Declare PtrSafe Function GetWindowText Lib "user32" Alias "GetWindowTextA" _
                                                            (ByVal hwnd As Long, _
                                                            ByVal lpString As String, _
                                                            ByVal cch As Long) As Long
#Else
' https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-findwindowa
Private Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" _
                                                        (ByVal lpClassName As String, _
                                                        ByVal lpWindowName As String) As Long
' https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-findwindowexa
Private Declare Function FindWindowEx Lib "user32.dll" Alias "FindWindowExA" _
                                                            (ByVal hwndParent As Long, _
                                                            ByVal hwndChildAfter As Long, _
                                                            ByVal lpszClass As String, _
                                                            ByVal lpszWindow As String) As Long
' https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-sendmessagea
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
                                                            (ByVal hwnd As Long, _
                                                            ByVal wMsg As Long, _
                                                            ByVal wParam As Long, _
                                                            ByRef lParam As Any) As Long
'https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-postmessagea
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" _
                                                            (ByVal hwnd As Long, _
                                                            ByVal wMsg As Long, _
                                                            ByVal wParam As Long, _
                                                            ByRef lParam As Any) As Long
'https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-keybd_event
Private Declare Sub keybd_event Lib "user32.dll" _
                                                            (ByVal bVk As Byte, _
                                                            ByVal bScan As Byte, _
                                                            ByVal dwFlags As Long, _
                                                            ByVal dwExtraInfo As Long)
'https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-getkeystate
Private Declare Function GetKeyState Lib "user32" ( _
                                                            ByVal nVirtKey As Long) As Long
'https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-getwindow
Public Declare Function GetWindow _
                                                        Lib "user32" _
                                                        (ByVal hwnd As Long, _
                                                        ByVal wCmd As Long) As Long
'https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-getclassname
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" _
                                                            (ByVal hwnd As Long, _
                                                            ByVal lpClassName As String, _
                                                            ByVal nMaxCount As Long) As Long
'https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-getwindowtexta
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" _
                                                            (ByVal hwnd As Long, _
                                                            ByVal lpString As String, _
                                                            ByVal cch As Long) As Long
#End If

'######################################
' ���� ��� ����
'######################################
' ����Ű ASCII �ڵ�
Public Const WM_SETTEXT = &HC
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const VK_RETURN = &HD
Public Const VK_CONTROL = &H11
Public Const VK_ESC = &H1B
Public Const KEYEVENTF_KEYUP As Long = 2
Public Const VK_A = &H41
Public Const VK_V = &H56
Public Const VK_C = &H43
Public Const VK_DOWN = &H28
Public Const VK_UP = &H26

Sub Test()

'����������������������������������������������������������������������
SendKakao "�������(ģ���̸�)", "�����޽���"

End Sub

'######################################
' ģ������ īī���� �޼����� �����մϴ�.
'######################################
Function SendKakao(Target As String, Message As String, Optional ChatRoomSearch As Boolean = False, Optional iDelay As Long = 1) As Boolean

' ���� ����
Dim hwnd_RichEdit As Long       ' ä���Է�â hwnd

' ģ�� ä���Է�â hwnd ã��
hwnd_RichEdit = FindRecepientHwnd(Target, ChatRoomSearch, iDelay)
'����������������������������������������������������������������������

'���� ����� ä��â ���� ��� �ȳ��޼��� ��� �� ����
If hwnd_RichEdit = 0 Then MsgBox "���� ����� īī���� ��ȭâ�� ã�� �� �����ϴ�.": Exit Function

' �޼��� ������
Send_TextMsg Message, hwnd_RichEdit

End Function

'######################################
' īī���� �޼����� �����մϴ�.
'######################################
Sub Send_TextMsg(Message As String, hwnd_RichEdit As Long)

' ��ȭ��� ä�� �Է�â hWnd �� �޼��� �Է�
Call SendMessage(hwnd_RichEdit, WM_SETTEXT, 0, ByVal Message)
'����������������������������������������������������������������������
' ����� Ctrl Ű �Է¿��� Ȯ��
If IsCtrlKeyDown = False Then
    ' Ctrl Ű ���Է� ��, �޼��� ����
    Call PostMessage(hwnd_RichEdit, WM_KEYDOWN, VK_RETURN, 0)
Else
    ' Ctrl Ű �Է����� ���, ������ Ctrl Ű �ø� -> �޼��� ���� -> Ctrl Ű ���Է�
    keybd_event VK_CONTROL, 0, KEYEVENTF_KEYUP, 0
    Call PostMessage(hwnd_RichEdit, WM_KEYDOWN, VK_RETURN, 0)
    keybd_event VK_CONTROL, 0, 0, 0
End If
    
End Sub

'######################################
' ���� ����� hWnd ���� ã���ϴ�.
' ���� ����� ä��â�� ���� ��� False�� ��ȯ�մϴ�.
'######################################
Function FindRecepientHwnd(Target As String, ChatRoomSearch As Boolean, iDelay As Long) As Long

' ���� ����
Dim dStart As Double
Dim hwnd_KakaoTalk As Long   ' ģ�� ä��â hwnd
Dim hwnd_RichEdit As Long       ' ä���Է�â hwnd

' ��������� ī��â ����
ActiveChat Target, iDelay, ChatRoomSearch
'����������������������������������������������������������������������

' ��������� ī��â Hwnd ã��
hwnd_KakaoTalk = FindWindow(vbNullString, Target)
'����������������������������������������������������������������������

dStart = Now
While hwnd_KakaoTalk = 0
    hwnd_KakaoTalk = FindWindow(vbNullString, Target)
    ' 1�� ���������� â�� ��ã���� �Լ� ����� False ��ȯ �� ����
    If DateDiff("s", dStart, Now) > 1 Then FindRecepientHwnd = 0: Exit Function
Wend

' ģ���� ä���Է�â hWnd ã��
hwnd_RichEdit = FindWindowEx(hwnd_KakaoTalk, 0, "RichEdit50W", vbNullString)  'īī���� ������ ���� RichEdit ClassName ����
If hwnd_RichEdit = 0 Then hwnd_RichEdit = FindWindowEx(hwnd_KakaoTalk, 0, "RichEdit20W", vbNullString)

FindRecepientHwnd = hwnd_RichEdit

End Function

'############################
' ģ�� ä��â�� �������� ���� ��� ä��â�� �˻� �� Ȱ��ȭ�մϴ�.
'############################
Function ActiveChat(Target As String, iDelay As Long, ChatRoomSearch As Boolean)

' ���� ����
Dim hWndMain As Long: Dim hWndChild1 As Long: Dim hWndChild2 As Long: Dim hWndEdit As Long
Dim i As Long

' ��ȭ��� â �̹� �������� �� ��ɹ� ����
If FindWindow(vbNullString, Target) > 0 Then Exit Function
'����������������������������������������������������������������������

' īī���� ���� hWnd �˻� [����������F5Ű]
hWndMain = FindHwndEVA
' īī���� ���� â ������ ī�� ���� �ȵ� -> �Լ� False ��ȯ �� ��ɹ� ����
If hWndMain = 0 Then ActiveChat = False: Exit Function

' ���� [����������F5Ű]
hWndChild1 = FindWindowEx(hWndMain, 0, "EVA_ChildWindow", vbNullString)
'����������������������������������������������������������������������
' ����ó ��� �˻� -> EVA_Window ù��° �׸�
hWndChild2 = FindWindowEx(hWndChild1, 0, "EVA_Window", vbNullString)
' ä��â ��� �˻� -> EVA_Window �ι�° �׸�
If ChatRoomSearch = True Then hWndChild2 = FindWindowEx(hWndChild1, hWndChild2, "EVA_Window", vbNullString)
' EVA_Window�� ��ȭ��� �˻�â
hWndEdit = FindWindowEx(hWndChild2, 0, "Edit", vbNullString)
  
' �˻�â�� ��ȭ��� ����/�ٿ��ֱ�
Call SendMessage(hWndEdit, WM_SETTEXT, 0, ByVal Target): MyDelay iDelay
' ä��â ��� �˻��� ��� ���ʹ���Ű ������ ù��° �׸� Ȱ��ȭ
If ChatRoomSearch = True Then Call PostMessage(hWndEdit, WM_KEYDOWN, VK_UP, 0): MyDelay iDelay
' ����Ű�� ä��â ����
Call PostMessage(hWndEdit, WM_KEYDOWN, VK_RETURN, 0): MyDelay iDelay

End Function

'############################
' īī���� ���࿩�� Ȯ�� ��, �������� �� ����â�� hWnd �� ��ȯ�մϴ�.
'############################
Private Function FindHwndEVA() As Long

' ���� ����
Dim hwnd As Long: Dim lngT As Long: Dim strT As String

' ���� Desktop���� �������� ù��° ���α׷� hWnd
hwnd = FindWindowEx(0, 0, vbNullString, vbNullString)

' ��� hWnd �� ���ư��� �˻�
While hwnd <> 0
    ' hWnd�� ���ư��� ClassName�� Ȯ��
    strT = String(100, Chr(0))
    lngT = GetClassName(hwnd, strT, 100)
    ' hWnd �� ClassName �� "EVA_Window_DblClk"�� ���Ե� ���
    If InStr(1, Left(strT, lngT), "EVA_Window_Dblclk") > 0 Then
        ' �ش� hWnd�� ä��â�̸��� �޾ƿɴϴ�.
        strT = String(100, Chr(0))
        lngT = GetWindowText(hwnd, strT, 100)
        ' ä��â �̸��� "īī����" �Ǵ� "KakaoTalk"(����OS ����)�� ���, hWnd �� �Լ� ����� ��ȯ �� ����
        If InStr(1, Left(strT, lngT), "īī����") > 0 Or InStr(1, Left(strT, lngT), "KakaoTalk") > 0 Then FindHwndEVA = hwnd: Exit Function
    End If
hwnd = FindWindowEx(0, hwnd, vbNullString, vbNullString)
Wend

End Function

'###############################################################
'�����ο��� VBA ����������Լ� (https://www.oppadu.com)
'�� MyDelay �Լ�
'�� �ӽ� ���� �ڵ�
'�� ______________iDelay    : ���� �ӵ� ���� (Ŭ���� ���� ����, 1 �� 4ms)
'###############################################################
Sub MyDelay(Optional iDelay As Long = 1)

Dim i As Long
For i = 1 To iDelay * 10000000:        i = i + 1:        Next

End Sub

'###############################################################
'�����ο��� VBA ����������Լ� (https://www.oppadu.com)
'�� IsCtrlKeyDown �Լ�
'�� Ű���� Ctrl Ű �������θ� Ȯ���մϴ�.
'�� �μ� ����
'_____________LeftRightKey : ����/������ Ctrl Ű �������θ� ���մϴ�.
' 1 : ���� CtrlŰ �Է½� TRUE
' 2 : ������ CtrlŰ �Է½� TRUE
' 3 : ���� CtrlŰ ���� �Է½� TRUE
' 0 : �� �� �ϳ��� �Է½� TRUE
'###############################################################

Private Function IsCtrlKeyDown(Optional LeftRightKey As Long = 0) As Boolean

Const VK_LCTRL = &HA2: Const VK_RCTRL = &HA3: Const KEY_MASK As Integer = &HFF80

Dim Result As Long

Select Case LeftRightKey
    '���� CTRL Ű �Է¿��� Ȯ��
    Case 1:        Result = GetKeyState(VK_LCTRL) And KEY_MASK
    '������ CTRL Ű �Է¿��� Ȯ��
    Case 2:        Result = GetKeyState(VK_RCTRL) And KEY_MASK
    '���� CTRL Ű ���� �Է¿��� Ȯ��
    Case 3:        Result = GetKeyState(VK_LCTRL) And GetKeyState(VK_RCTRL) And KEY_MASK
    'CTRL Ű ���� �ϳ��� �Է¿��� Ȯ��
    Case Else:    Result = GetKeyState(vbKeyControl) And KEY_MASK
End Select

IsCtrlKeyDown = CBool(Result)

End Function


