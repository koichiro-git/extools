Attribute VB_Name = "mdlDatePicker"
'// ////////////////////////////////////////////////////////////////////////////
'// �v���W�F�N�g   : �g���c�[��
'// �^�C�g��       : ���t�s�b�J�[�p�W�����W���[��
'// ���W���[��     : mdlDatePicker
'// ����           : ���t�s�b�J�[�t�H�[���Ŏg�p����E�B���h�E�v���V�[�W�����L��
'// ////////////////////////////////////////////////////////////////////////////
'// Copyright (c) by Koichiro.
'// ////////////////////////////////////////////////////////////////////////////
Option Explicit
Option Base 0

'// ////////////////////////////////////////////////////////////////////////////
'// Windows API �֘A�̐錾
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
Public Declare PtrSafe Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As LongPtr, ByVal hwnd As LongPtr, ByVal Msg As LongPtr, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr


Private Const WM_NOTIFY = &H4E           '//0x004E
Private Const MCN_SELECT = -746


'// ���t�����iMonthView ����MCM_GETCURSEL�w��œ��t���擾����ۂɎg�p�j
Private Type SYSTEMTIME
    wYear           As Integer
    wMonth          As Integer
    wDayOfWeek      As Integer
    wDay            As Integer
    wHour           As Integer
    wMinute         As Integer
    wSecond         As Integer
    wMilliseconds   As Integer
End Type

Private Type NMHDR
    hwndFrom        As LongPtr
    idFrom          As LongPtr
    code            As Long  'Integer
End Type

Private Type tagNMSELCHANGE
    hdr             As NMHDR
    stSelStart      As SYSTEMTIME
    stSelEnd        As SYSTEMTIME
End Type


Public defaultProcAddress   As LongPtr
Public hMonthView           As LongPtr  '// MonthView�̃E�B���h�E�n���h��


'// ////////////////////////////////////////////////////////////////////////////
'// ���\�b�h�F   ���t�s�b�J�[�p�E�B���h�E�v���V�[�W��
'// �����F       ���t�I�����̏���
'// �����F       hwnd:   �E�B���h�E�n���h��
'//              uMsg:   ���b�Z�[�W
'//              wParam: �ǉ��̃��b�Z�[�W�ŗL���
'//              lParam: �ǉ��̃��b�Z�[�W�ŗL���
'// ////////////////////////////////////////////////////////////////////////////
Public Function WindowProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As LongPtr
On Error GoTo ErrorHandler
    Dim tagNMHDR    As NMHDR
    Dim prm         As tagNMSELCHANGE
    
    '// ToDo:�@�u�b�N�������Ƃ��̏���
    If uMsg = WM_NOTIFY Then
        Call CopyMemory(tagNMHDR, ByVal lParam, Len(tagNMHDR))
        If tagNMHDR.hwndFrom = hMonthView And tagNMHDR.code = MCN_SELECT Then
            Call CopyMemory(prm, ByVal lParam, Len(prm))
            ActiveCell.Value = CDate(prm.stSelStart.wYear & "/" & prm.stSelStart.wMonth & "/" & prm.stSelStart.wDay)
        End If
    End If
    '// �����̐��ۂɂ�����炸�K���f�t�H���g�̃E�B���h�E�v���V�[�W���Ɉ����n��
ErrorHandler:
    WindowProc = CallWindowProc(defaultProcAddress, hwnd, uMsg, wParam, lParam)
End Function
