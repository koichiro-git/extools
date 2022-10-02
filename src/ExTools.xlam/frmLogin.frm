VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmLogin 
   Caption         =   "���O�C��"
   ClientHeight    =   3255
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3510
   OleObjectBlob   =   "frmLogin.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'// ////////////////////////////////////////////////////////////////////////////
'// �v���W�F�N�g   : �g���c�[��
'// �^�C�g��       : �F�؉��
'// ���W���[��     : frmLogin
'// ����           : �ڑ���ւ̔F�؂��s���B
'// ////////////////////////////////////////////////////////////////////////////
'// Copyright (c) by Koichiro.
'// ////////////////////////////////////////////////////////////////////////////
Option Explicit
Option Base 0

'// ////////////////////////////////////////////////////////////////////////////
'// �v���C�x�[�g�萔
'// �Ȃ�


'// ////////////////////////////////////////////////////////////////////////////
'// �C�x���g�F �t�H�[�� ��������
Private Sub UserForm_Initialize()
    '// �R���{�{�b�N�X�ݒ�
    Call gsSetCombo(cmbMethod, dct_excel & ",Excel;" & dct_odbc & ",ODBC", 0)
    
    '// �L���v�V�����ݒ�
    frmLogin.Caption = LBL_LGI_FORM
    lblUid.Caption = LBL_LGI_UID
    lblPassword.Caption = LBL_LGI_PASSWORD
    lblConnStr.Caption = LBL_LGI_STRING
    lblConnTo.Caption = LBL_LGI_CONN_TO
    cmdOk.Caption = LBL_LGI_LOGIN
    cmdCancel.Caption = LBL_LGI_CANCEL
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// �C�x���g�F �t�H�[�� �A�N�e�B�u��
Private Sub UserForm_Activate()
    If cmbMethod.Value <> dct_excel Then
        '// ���[�UID�܂��̓p�X���[�h�Ƀt�H�[�J�X��ݒ�
        If txtUserId.Text = BLANK Then
            Call txtUserId.SetFocus
        Else
            Call txtPassword.SetFocus
        End If
    End If
    '// �p�X���[�h�͏�ɋ�
    txtPassword.Text = BLANK
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// �C�x���g�F �n�j�{�^�� �N���b�N��
Private Sub cmdOk_Click()
On Error GoTo ErrorHandler
    Dim isConnected As Boolean
    
    '// �����̃R�l�N�V������ؒf����
    If Not gADO Is Nothing Then
        Call gADO.CloseConnection
    End If
    Set gADO = Nothing
  
    '// ���̓`�F�b�N
    If (cmbMethod.Value = dct_excel) Then     '// Excel�ւ̐ڑ��ł͌��݂̃t�@�C���`�F�b�N
        If (ActiveWorkbook.Path = BLANK) Or (Not ActiveWorkbook.Saved) Then
            If MsgBox(MSG_NEED_EXCEL_SAVED, vbYesNo, APP_TITLE) = vbNo Then
                cmbMethod.SetFocus
                Exit Sub
            Else
                Call ActiveWorkbook.Save
            End If
        End If
    ElseIf (txtUserId.Text = BLANK) Then      '// ���[�UID
        Call MsgBox(MSG_NEED_FILL_ID, vbOKOnly, APP_TITLE)
        Call txtUserId.SetFocus
        Exit Sub
    ElseIf (txtPassword.Text = BLANK) Then    '// �p�X���[�h
        Call MsgBox(MSG_NEED_FILL_PWD, vbOKOnly, APP_TITLE)
        Call txtPassword.SetFocus
        Exit Sub
    ElseIf (txtHostString.Text = BLANK) Then  '// �ڑ�������
        Call MsgBox(MSG_NEED_FILL_TNS, vbOKOnly, APP_TITLE)
        Call txtHostString.SetFocus
        Exit Sub
    End If
  
    '// �ڑ�
    Set gADO = New cADO
    Select Case cmbMethod.Value
'        Case dct_oracle  '// v2���I���N���͑ΏۊO
'            isConnected = gOra.Initialize(txtHostString.Text, txtUserId.Text, txtPassword.Text, dct_oracle)
        Case dct_odbc
            isConnected = gADO.Initialize(txtHostString.Text, txtUserId.Text, txtPassword.Text, dct_odbc)
        Case dct_excel
            isConnected = gADO.Initialize(ActiveWorkbook.Path & "\" & ActiveWorkbook.Name, BLANK, BLANK, dct_excel)
    End Select
  
    '// ����
    If Not isConnected Then
        Call MsgBox(MSG_LOG_ON_FAILED & vbLf & vbLf _
                & "Error Number: " & gADO.NativeError & vbLf _
                & "Error Description: " & gADO.ErrorText, vbOKOnly, APP_TITLE)
    Else
        Call Me.Hide
        Call MsgBox(MSG_LOG_ON_SUCCESS, vbOKOnly, APP_TITLE)
    End If
    Exit Sub
  
ErrorHandler:
    Call gsShowErrorMsgDlg("frmLogin.cmdOk_Click", Err)
    Set gADO = Nothing
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// �C�x���g�F �L�����Z���{�^�� �N���b�N��
Private Sub cmdCancel_Click()
    Call Me.Hide
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// �C�x���g�F �ڑ���R���{ �ύX��
Private Sub cmbMethod_Change()
    '// �G�N�Z���ڑ����́A���[�UID�A�p�X���[�h�A�ڑ�������͎g�p�s�ɐݒ�
    If cmbMethod.Value = dct_excel Then
        txtUserId.Enabled = False
        txtUserId.BackColor = CLR_DISABLED
        txtPassword.Enabled = False
        txtPassword.BackColor = CLR_DISABLED
        txtHostString.Enabled = False
        txtHostString.BackColor = CLR_DISABLED
    Else
        txtUserId.Enabled = True
        txtUserId.BackColor = CLR_ENABLED
        txtPassword.Enabled = True
        txtPassword.BackColor = CLR_ENABLED
        txtHostString.Enabled = True
        txtHostString.BackColor = CLR_ENABLED
    End If
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// END
'// ////////////////////////////////////////////////////////////////////////////
