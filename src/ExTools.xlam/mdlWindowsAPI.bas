Attribute VB_Name = "mdlWindowsAPI"
'// ////////////////////////////////////////////////////////////////////////////
'// �v���W�F�N�g   : �g���c�[��
'// �^�C�g��       : Windows API�錾
'// ���W���[��     : mdlWindowsAPI
'// ����           : Windows API �֘A�̐錾�BWin32API_PtrSafe.txt�x�[�X
'// ////////////////////////////////////////////////////////////////////////////
'// Copyright (c) by Koichiro.
'// ////////////////////////////////////////////////////////////////////////////
Option Explicit
Option Base 0

'// ////////////////////////////////////////////////////////////////////////////
'// �萔

'// Window Styles
Public Const WS_VISIBLE = &H10000000
Public Const WS_CHILD = &H40000000
'Private Const WS_BORDER = &H800000
Public Const WS_EX_TOOLWINDOW = &H80

'// Window field offsets for GetWindowLong() and GetWindowWord()
Public Const GWL_WNDPROC = (-4)
Public Const GWL_EXSTYLE = (-20)

'// MonthView SendMessage�p���b�Z�[�W��`
Public Const MCM_FIRST = &H1000
Public Const MCM_GETCURSEL = (MCM_FIRST + 1)           '// �I�����ꂽ���t���擾
Public Const MCM_SETCURSEL = (MCM_FIRST + 2)
Public Const MCM_GETMINREQRECT = (MCM_FIRST + 9)       '// MonthView�̃T�C�Y���擾

'// MonthView�֘A
Public Const MONTHCAL_CLASS = "SysMonthCal32"
Public Const ICC_DATE_CLASSES = &H100
Public Const WM_NOTIFY = &H4E           '//0x004E
Public Const MCN_SELECT = -746

'// �t�H���_�I���_�C�A���O�֘A
Public Const BIF_RETURNONLYFSDIRS = &H1
Public Const MAX_PATH = 260


'// ////////////////////////////////////////////////////////////////////////////
'// �^�C�v

'// InitCommonControlsEx�p
Public Type tagINITCOMMONCONTROLSEX
    dwSize          As Long
    dwICC           As Long
End Type

'// ���t�����iMonthView ����MCM_GETCURSEL�w��œ��t���擾����ۂɎg�p�j
Public Type SYSTEMTIME
    wYear           As Integer
    wMonth          As Integer
    wDayOfWeek      As Integer
    wDay            As Integer
    wHour           As Integer
    wMinute         As Integer
    wSecond         As Integer
    wMilliseconds   As Integer
End Type

'// RECT�i�E�B���h�E�T�C�Y�ݒ莞�Ɏg�p�B�P�ʁ��s�N�Z���j
Public Type RECT
    Left            As Long
    Top             As Long
    Right           As Long
    Bottom          As Long
End Type

'// ���W �i�}�E�X���W�j
Public Type POINTAPI
    x               As Long
    y               As Long
End Type

'// �t�H���_�I���_�C�A���O�֘A
Public Type BROWSEINFO
    hwndOwner       As LongPtr
    pidlRoot        As Long
    pszDisplayName  As String
    lpszTitle       As String
    ulFlags         As Long
    lpfn            As LongPtr
    lParam          As LongPtr
    iImage          As Long
End Type

'// �E�C���h�E���b�Z�[�W�Ǘ��֘A
Public Type NMHDR
    hwndFrom        As LongPtr
    idFrom          As LongPtr
    code            As Long
End Type

Public Type tagNMSELCHANGE
    hdr             As NMHDR
    stSelStart      As SYSTEMTIME
    stSelEnd        As SYSTEMTIME
End Type


'// ////////////////////////////////////////////////////////////////////////////
'// �֐�

'// �E�C���h�E����
Public Declare PtrSafe Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As LongPtr, ByVal hMenu As LongPtr, ByVal hInstance As LongPtr, lpParam As Any) As LongPtr
'// �E�C���h�E���W�ݒ�
Public Declare PtrSafe Function MoveWindow Lib "user32" (ByVal hWnd As LongPtr, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
'// �}�E�X���W�擾
Public Declare PtrSafe Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
'// �E�B���h�E�R���g���[���̑���
Public Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, lParam As Any) As LongPtr
'// �E�B���h�E�n���h���擾
Public Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
Public Declare PtrSafe Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As LongPtr, ByVal hWnd2 As LongPtr, ByVal lpsz1 As String, ByVal lpsz2 As String) As LongPtr

'// �R�����R���g���[��������
Public Declare PtrSafe Function InitCommonControlsEx Lib "ComCtl32" (LPINITCOMMONCONTROLSEX As Any) As Long

'// �E�B���h�E�X�^�C���␳
#If Win64 Then
    Public Declare PtrSafe Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongPtrA" (ByVal hWnd As LongPtr, ByVal nIndex As Long) As LongPtr
    Public Declare PtrSafe Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongPtrA" (ByVal hWnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
#Else
    Public Declare PtrSafe Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As LongPtr, ByVal nIndex As Long) As LongPtr
    Public Declare PtrSafe Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
#End If

'// �������R�s�[�i�L���X�g��ցj
Public Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
'// �E�B���h�E �v���V�[�W���Ăяo��
Public Declare PtrSafe Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As LongPtr, ByVal hWnd As LongPtr, ByVal Msg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr

'// �t�H���_�I��
Public Declare PtrSafe Function apiSHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolder" (lpBrowseInfo As BROWSEINFO) As LongPtr
'// �p�X�擾
Public Declare PtrSafe Function apiSHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDList" (ByVal piDL As LongPtr, ByVal strPath As String) As LongPtr
'//�L�[���荞��
Public Declare PtrSafe Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
