VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAbout 
   Caption         =   "�G�N�Z���g���c�[��"
   ClientHeight    =   2805
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6030
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
    Me.Caption = APP_TITLE
    lblVersion.Caption = APP_TITLE & Space(1) & APP_VERSION
    lblRepo.Caption = LBL_ABT_REPO
    lblManual.Caption = LBL_ABT_MANUAL
    cmdClose.Caption = LBL_COM_CLOSE
End Sub


'// //////////////////////////////////////////////////////////////////
'// �C�x���g�F ����{�^�� �N���b�N��
Private Sub cmdClose_Click()
  Call Me.Hide
End Sub


'// //////////////////////////////////////////////////////////////////
'// �C�x���g�F ���x��_�_�E�����[�h �N���b�N��
Private Sub lblDownload_Click()
    Call ThisWorkbook.FollowHyperlink(Address:="https://github.com/koichiro-git/extools/releases")
End Sub


'// //////////////////////////////////////////////////////////////////
'// �C�x���g�F ���x��_�}�j���A�� �N���b�N��
Private Sub lblUserManual_Click()
    Call ThisWorkbook.FollowHyperlink(Address:="https://koichiro-git.github.io/extools/")
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// END
'// ////////////////////////////////////////////////////////////////////////////
