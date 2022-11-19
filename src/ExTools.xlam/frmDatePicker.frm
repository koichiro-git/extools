VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDatePicker 
   Caption         =   "���t"
   ClientHeight    =   4020
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3525
   OleObjectBlob   =   "frmDatePicker.frx":0000
   ShowModal       =   0   'False
End
Attribute VB_Name = "frmDatePicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'// ////////////////////////////////////////////////////////////////////////////
'// �v���W�F�N�g   : �g���c�[��
'// �^�C�g��       : ���t�s�b�N�J�����_�[
'// ���W���[��     : frmDatePicker
'// ����           : MonthView���g�p����DatePicker
'// ////////////////////////////////////////////////////////////////////////////
'// Copyright (c) by Koichiro.
'// ////////////////////////////////////////////////////////////////////////////
Option Explicit
Option Base 0

'// ////////////////////////////////////////////////////////////////////////////
'// �A�v���P�[�V�����萔

Private Const THUNDER_FRAME     As String = "ThunderDFrame" '// Excel VBA���[�U�[�t�H�[���̃N���X���iExcel2000�ȍ~=ThunerDFrame / ����ȑO=ThunderXFrame�j

'// �_���C���`������̉�ʂ̃s�N�Z�����i�|�C���g���s�N�Z�����Z�W���j
'// GetDeviceCaps �����96��Ԃ����ߤ�v���O�����ł̓��I�擾����߁A�萔�Ƃ���iSetProcessDPIAware�͎������Ȃ��j
'// https://learn.microsoft.com/ja-jp/windows-hardware/manufacture/desktop/dpi-related-apis-and-registry-settings?view=windows-11
Private Const LOG_PIXELS        As Long = 96


'// ////////////////////////////////////////////////////////////////////////////
'// Windows API �֘A�̐錾

'// �萔
Private Const MONTHCAL_CLASS = "SysMonthCal32"
Private Const ICC_DATE_CLASSES = &H100          '// �R�����R���g���[���p�萔�i���t�Ǝ����̑I���R���g���[���j

'// Window Styles
Private Const WS_VISIBLE = &H10000000
Private Const WS_CHILD = &H40000000
Private Const WS_EX_TOOLWINDOW = &H80

'// Window field offsets for GetWindowLong() and GetWindowWord()
Const GWL_EXSTYLE = (-20)

'// MonthView SendMessage�p���b�Z�[�W��`
Private Const MCM_FIRST = &H1000
Private Const MCM_GETCURSEL = (MCM_FIRST + 1)           '// �I�����ꂽ���t���擾
Private Const MCM_SETCURSEL = (MCM_FIRST + 2)
Private Const MCM_GETMINREQRECT = (MCM_FIRST + 9)       '// MonthView�̃T�C�Y���擾

'// �^�C�v
'// InitCommonControlsEx�p
Private Type tagINITCOMMONCONTROLSEX
    dwSize          As Long
    dwICC           As Long
End Type

'// ���t�����iMonthView ����MCM_GETCURSEL�w��œ��t���擾����ۂɎg�p�j
Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

'// RECT�i�E�B���h�E�T�C�Y�ݒ莞�Ɏg�p�B�P�ʁ��s�N�Z���j
Private Type RECT
    Left        As Long
    Top         As Long
    Right       As Long
    Bottom      As Long
End Type

'// ���W �i�}�E�X���W�j
Private Type POINTAPI
    x           As Long
    y           As Long
End Type

'// �E�C���h�E����
Private Declare PtrSafe Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As LongPtr, ByVal hMenu As LongPtr, ByVal hInstance As LongPtr, lpParam As Any) As LongPtr
'// �E�C���h�E���W�ݒ�
Private Declare PtrSafe Function MoveWindow Lib "user32" (ByVal hwnd As LongPtr, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
'// �}�E�X���W�擾
Private Declare PtrSafe Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
'// �E�B���h�E�R���g���[���̑���
Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, lParam As Any) As LongPtr
'// �E�B���h�E���ݔ���
Private Declare PtrSafe Function IsWindow Lib "user32" (ByVal hwnd As LongPtr) As Long
'// �E�B���h�E�p���i��d�N�����j
Private Declare PtrSafe Function DestroyWindow Lib "user32" (ByVal hwnd As LongPtr) As Long
'// �E�B���h�E�n���h���擾
Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
Private Declare PtrSafe Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As LongPtr, ByVal hWnd2 As LongPtr, ByVal lpsz1 As String, ByVal lpsz2 As String) As LongPtr
'// �R�����R���g���[��������
Private Declare PtrSafe Function InitCommonControlsEx Lib "ComCtl32" (LPINITCOMMONCONTROLSEX As Any) As Long

'// �E�B���h�E�X�^�C���␳
#If Win64 Then
    Private Declare PtrSafe Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongPtrA" (ByVal hwnd As LongPtr, ByVal nIndex As Long) As LongPtr
    Private Declare PtrSafe Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongPtrA" (ByVal hwnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
#Else
    Private Declare PtrSafe Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As LongPtr, ByVal nIndex As Long) As LongPtr
    Private Declare PtrSafe Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
#End If
    
'// ////////////////////////////////////////////////////////////////////////////
'// �ϐ�
Private hwndMonthView           As LongPtr  '// MonthView�̃E�B���h�E�n���h��


'// //////////////////////////////////////////////////////////////////
'// �C�x���g�F �t�H�[�� ��������
Private Sub UserForm_Initialize()
    Call psSetupDatePicker
End Sub


'// //////////////////////////////////////////////////////////////////
'// �C�x���g�F OK�{�^�� �N���b�N��
Private Sub cmdExecute_Click()
    Dim st          As SYSTEMTIME
    
    Call SendMessage(hwndMonthView, MCM_GETCURSEL, 0, st)
    
    '// �o��(�I��͈͂��Z���̏ꍇ�̂�)
    If TypeName(Selection) = TYPE_RANGE Then
        ActiveCell.Value = CDate(st.wYear & "/" & st.wMonth & "/" & st.wDay)
    End If
    
    '// �u��ɊJ���v�g�O�����j���[���N���b�N��ԂłȂ���΁A�{��ʂ����
    If gDatePickerToggle Then
        Call Unload(Me)
    End If
End Sub


Private Sub psSetupDatePicker()
'On Error GoTo ErrorHandler
    Dim icce                As tagINITCOMMONCONTROLSEX
    Dim rc                  As RECT
    Dim lnghWnd_Sub         As LongPtr
    Dim hwndForm            As LongPtr  '// UserForm�̃E�B���h�E�n���h��
    Dim lResult             As LongPtr
    
    ' �R�����R���g���[��������
    icce.dwICC = ICC_DATE_CLASSES
    icce.dwSize = Len(icce)
    lResult = InitCommonControlsEx(icce)
    If lResult = 0 Then Call Err.Raise(Number:=513, Description:="���t�s�b�J�[��ʂ𐶐��ł��܂���")
    
    ' ���[�U�[�t�H�[����HWND�̎擾
    hwndForm = FindWindow(THUNDER_FRAME, Me.Caption)
    If hwndForm = 0 Then Call Err.Raise(Number:=513, Description:="���t�s�b�J�[��ʂ𐶐��ł��܂���")

    ' MonthView�z�u�p�n���h���̎擾
    lnghWnd_Sub = FindWindowEx(hwndForm, 0, vbNullString, vbNullString)
    
    '// MonthView�E�B���h�E����(�T�C�Y�[���Ő����@https://learn.microsoft.com/ja-jp/windows/win32/controls/mcm-getminreqrect)
    hwndMonthView = CreateWindowEx(0, MONTHCAL_CLASS, vbNullString, (WS_VISIBLE Or WS_CHILD), 0, 0, 0, 0, lnghWnd_Sub, 0, 0, vbNullString) '//lnghWnd_Sub, 0, lnghInstance, vbNullString)
    '// MonthView�p�E�B���h�E�̃��T�C�Y
    lResult = SendMessage(hwndMonthView, MCM_GETMINREQRECT, 0, rc)
    Call MoveWindow(hwndMonthView, 0, 0, rc.Right - rc.Left, rc.Bottom - rc.Top, 1&)

    '// ��ʕ␳ //////////
    Call SetWindowLongPtr(hwndForm, GWL_EXSTYLE, GetWindowLongPtr(hwndForm, GWL_EXSTYLE) Or WS_EX_TOOLWINDOW)   '// UserForm��ToolWindow�X�^�C���ɕύX
    
    '// �t�H�[���T�C�Y�␳�B�J�����_�[��(px��pt�ϊ�)�{�t�H�[���g���B�c���̓J�����_�[��+�{�^����+�X�y�[�T�[1pt
    Me.Width = (rc.Right - rc.Left) * 72 / LOG_PIXELS + (Me.Width - Me.InsideWidth)
    Me.Height = rc.Bottom * 72 / LOG_PIXELS + cmdExecute.Height + 2 + (Me.Height - Me.InsideHeight)
    
    '// OK�{�^���̃T�C�Y�E�ʒu�␳�iMonthView�̉��B�t�H�[���T�C�Y�ɍ��킹�ĉ�����ݒ�j
    cmdExecute.Width = (Me.InsideWidth - 2)
    cmdExecute.Left = 1
    cmdExecute.Top = rc.Bottom * 72 / LOG_PIXELS + 1
    
    '// �t�H�[���ʒu�␳�i�}�E�X���W�ցj
    Call MoveFormToMouse
    Exit Sub

ErrorHandler:
    '// none
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// ���\�b�h�F    �t�H�[���ʒu�␳
'// �����F        �t�H�[�����}�E�X�ʒu�Ɉړ�������
'// �����F        �Ȃ�
'// ////////////////////////////////////////////////////////////////////////////
Private Sub MoveFormToMouse()
    Dim mousePos As POINTAPI
    
    Call GetCursorPos(mousePos)                 '// �}�E�X�ʒu�擾
    Me.Left = 72 / LOG_PIXELS * mousePos.x
    Me.Top = 72 / LOG_PIXELS * mousePos.y
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// END.
'// ////////////////////////////////////////////////////////////////////////////
