VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAbout 
   Caption         =   "�G�N�Z���g���c�[��"
   ClientHeight    =   2535
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6120
   OleObjectBlob   =   "frmAbout.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'// ////////////////////////////////////////////////////////////////////////////
'// �v���W�F�N�g   : �g���c�[��
'// �^�C�g��       : About��ʃt�H�[��
'// ���W���[��     : frmAbout
'// ����           : �o�[�W�������A���|�W�g�������o�͂���
'// ////////////////////////////////////////////////////////////////////////////
'// Copyright (c) by Koichiro.
'// ////////////////////////////////////////////////////////////////////////////
Option Explicit
Option Base 0


'// //////////////////////////////////////////////////////////////////
'// �C�x���g�F �t�H�[�� ��������
Private Sub UserForm_Initialize()
    lblVersion = APP_TITLE & Space(1) & APP_VERSION

End Sub


'// //////////////////////////////////////////////////////////////////
'// �C�x���g�F ����{�^�� �N���b�N��
Private Sub cmdClose_Click()
  Call Me.Hide
End Sub


'// //////////////////////////////////////////////////////////////////
'// �C�x���g�F ���x��_�_�E�����[�h �N���b�N��
Private Sub lblDownload_Click()
    ThisWorkbook.FollowHyperlink Address:="https://github.com/koichiro-git/extools/blob/main/bin/ExTools.xlam"
End Sub


'// //////////////////////////////////////////////////////////////////
'// �C�x���g�F ���x��_�}�j���A�� �N���b�N��
Private Sub lblUserManual_Click()
    ThisWorkbook.FollowHyperlink Address:="https://koichiro-git.github.io/extools/"
End Sub

