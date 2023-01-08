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
'// �ϐ�
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
Public Function WindowProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As LongPtr
On Error GoTo ErrorHandler
    Dim tagNMHDR    As NMHDR
    Dim prm         As tagNMSELCHANGE
    
    If uMsg = WM_NOTIFY Then
        Call CopyMemory(tagNMHDR, ByVal lParam, Len(tagNMHDR))
        If tagNMHDR.hwndFrom = hMonthView And tagNMHDR.code = MCN_SELECT Then
            Call CopyMemory(prm, ByVal lParam, Len(prm))
            ActiveCell.Value = CDate(prm.stSelStart.wYear & "/" & prm.stSelStart.wMonth & "/" & prm.stSelStart.wDay)
        End If
    End If
    '// �A�N�e�B�u�Z���������ꍇ�Ȃǂ��G���[�͂��ׂĖ�������邽�ߏ�������͂��Ȃ��B
    '// �����̐��ۂɂ�����炸�K���f�t�H���g�̃E�B���h�E�v���V�[�W���Ɉ����n�����߁AExit�͖���
    
ErrorHandler:
    WindowProc = CallWindowProc(defaultProcAddress, hWnd, uMsg, wParam, lParam)
End Function
