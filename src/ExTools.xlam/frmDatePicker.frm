VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDatePicker 
   Caption         =   "���t"
   ClientHeight    =   6060
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13890
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

Private Const CALENDAR_SEP_WIDTH    As Double = 6  '// �J�����_�[2�i2�����j���̊Ԋu4.5pt + �\��


'// //////////////////////////////////////////////////////////////////
'// �C�x���g�F �t�H�[�� ��������
Private Sub UserForm_Initialize()
    Call psSetupDatePicker
End Sub


Private Sub psSetupDatePicker()
On Error GoTo ErrorHandler
    Dim icce                As tagINITCOMMONCONTROLSEX
    Dim rc                  As RECT
    Dim hWnd_Sub            As LongPtr
    Dim hWnd                As LongPtr  '// UserForm�̃E�B���h�E�n���h��
    Dim lResult             As LongPtr
    Dim calendarWidth       As Long
    
    ' �R�����R���g���[��������
    icce.dwICC = ICC_DATE_CLASSES
    icce.dwSize = Len(icce)
    lResult = InitCommonControlsEx(icce)
    If lResult = 0 Then Call Err.Raise(Number:=513, Description:="���t�s�b�J�[��ʂ𐶐��ł��܂���")
    
    ' ���[�U�[�t�H�[����HWND�̎擾
    hWnd = FindWindow(THUNDER_FRAME, Me.Caption)
    If hWnd = 0 Then Call Err.Raise(Number:=513, Description:="���t�s�b�J�[��ʂ𐶐��ł��܂���")

    ' MonthView�z�u�p�n���h���̎擾
    hWnd_Sub = FindWindowEx(hWnd, 0, vbNullString, vbNullString)
    
    '// MonthView�E�B���h�E����(�T�C�Y�[���Ő����@https://learn.microsoft.com/ja-jp/windows/win32/controls/mcm-getminreqrect)
    hMonthView = CreateWindowEx(0, MONTHCAL_CLASS, vbNullString, (WS_VISIBLE Or WS_CHILD), 0, 0, 0, 0, hWnd_Sub, 0, 0, vbNullString) '//hWnd_Sub, 0, lnghInstance, vbNullString)
    '// MonthView�p�E�B���h�E�̃��T�C�Y
    lResult = SendMessage(hMonthView, MCM_GETMINREQRECT, 0, rc)
    calendarWidth = (rc.Right - rc.Left) * 2 + CALENDAR_SEP_WIDTH
    Call MoveWindow(hMonthView, 0, 0, calendarWidth, rc.Bottom - rc.Top, 1&)

    defaultProcAddress = SetWindowLongPtr(hWnd_Sub, GWL_WNDPROC, AddressOf WindowProc)

    '// ��ʕ␳ //////////
    Call SetWindowLongPtr(hWnd, GWL_EXSTYLE, GetWindowLongPtr(hWnd, GWL_EXSTYLE) Or WS_EX_TOOLWINDOW)   '// UserForm��ToolWindow�X�^�C���ɕύX
    
    '// �t�H�[���T�C�Y�␳(px��pt�ϊ�)�B�����̓J�����_�[���{�t�H�[���g���B�c���̓J�����_�[���{�t�H�[���g��
    Me.Width = calendarWidth * 72 / LOG_PIXELS + (Me.Width - Me.InsideWidth)
    Me.Height = rc.Bottom * 72 / LOG_PIXELS + (Me.Height - Me.InsideHeight)
    
    '// �t�H�[���ʒu�␳�i�}�E�X���W�ցj
    Call MoveFormToMouse
    Exit Sub

ErrorHandler:
    Call gsResumeAppEvents
    Call gsShowErrorMsgDlg("frmDatePicker.psSetupDatePicker", Err)
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// ���\�b�h�F    �t�H�[���ʒu�␳
'// �����F        �t�H�[�����}�E�X�ʒu�Ɉړ�������
'// �����F        �Ȃ�
'// ////////////////////////////////////////////////////////////////////////////
Private Sub MoveFormToMouse()
    Dim mousePos As POINTAPI
    
    Call GetCursorPos(mousePos)
    Me.Left = 72 / LOG_PIXELS * mousePos.x
    Me.Top = 72 / LOG_PIXELS * mousePos.y
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// END.
'// ////////////////////////////////////////////////////////////////////////////
