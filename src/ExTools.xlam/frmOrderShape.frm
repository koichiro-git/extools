VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmOrderShape 
   Caption         =   "�V�F�C�v�̔z�u"
   ClientHeight    =   3330
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4305
   OleObjectBlob   =   "frmOrderShape.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "frmOrderShape"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'// ////////////////////////////////////////////////////////////////////////////
'// �v���W�F�N�g   : �g���c�[��
'// �^�C�g��       : �V�F�C�v�̐���t�H�[��
'// ���W���[��     : frmOrderShape
'// ����           : �I�����ꂽ�V�F�C�v���Z���ɂ��킹�Đ��񂳂���
'// ////////////////////////////////////////////////////////////////////////////
'// Copyright (c) by Koichiro.
'// ////////////////////////////////////////////////////////////////////////////
Option Explicit
Option Base 0


'// //////////////////////////////////////////////////////////////////
'// �C�x���g�F �t�H�[�� ��������
Private Sub UserForm_Initialize()
    Dim idx   As Integer
    
    '// �R���{�{�b�N�X�ݒ�
    Call gsSetCombo(cmbHeight, CMB_ORD_HEIGHT, 0)
    Call gsSetCombo(cmbWidth, CMB_ORD_WIDTH, 0)
    
    With cmbMargin
        Call .Clear
        For idx = 0 To 8
            Call .AddItem(CStr(idx))
            .List(idx, 1) = CStr(idx) & " pt"
        Next
        .ListIndex = 1
    End With
    
    Call ckbDetail_Change   '// �`�F�b�N�{�b�N�X
    
    '// �L���v�V�����ݒ�
    frmOrderShape.Caption = LBL_ORD_FORM
    cmdExecute.Caption = LBL_COM_EXEC
    cmdClose.Caption = LBL_COM_CLOSE
    ckbDetail.Caption = LBL_ORD_OPTIONS
    lblMargin.Caption = LBL_ORD_MARGIN
    lblHeight.Caption = LBL_ORD_HEIGHT
    lblWidth.Caption = LBL_ORD_WIDTH
End Sub


'// //////////////////////////////////////////////////////////////////
'// �C�x���g�F ����{�^�� �N���b�N��
Private Sub cmdClose_Click()
    Call Me.Hide
End Sub


'// //////////////////////////////////////////////////////////////////
'// �C�x���g�F �ڍאݒ�`�F�b�N�{�b�N�X �ύX��
Private Sub ckbDetail_Change()
    '// �Z���Ƀt�B�b�g������Ƃ��̂ݗL��
    cmbHeight.Enabled = ckbDetail.Value
    cmbWidth.Enabled = ckbDetail.Value
End Sub


'// //////////////////////////////////////////////////////////////////
'// �C�x���g�F ���s�{�^�� �N���b�N��
Private Sub cmdExecute_Click()
On Error GoTo ErrorHandler
    Dim idx   As Integer
    
    For idx = 1 To ActiveWindow.Selection.ShapeRange.Count  '// shaperange�̊J�n�C���f�b�N�X�͂P����
        Call psSetShapePos(ActiveWindow.Selection.ShapeRange(idx), cmbMargin.Value)
    Next
    Exit Sub
  
ErrorHandler:
    Call MsgBox(MSG_SHAPE_NOT_SELECTED, vbOKOnly, APP_TITLE)
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// ���\�b�h�F   �V�F�C�v�ʒu�E�T�C�Y�ݒ�
'// �����F
'// �����F       targetShape: �ΏۃV�F�C�v�I�u�W�F�N�g
'//              ptMargin: �}�[�W��
'// ////////////////////////////////////////////////////////////////////////////
Private Sub psSetShapePos(targetShape As Shape, ptMargin As Integer)
    Dim basePos(4) As Long
    
    basePos(0) = targetShape.TopLeftCell.Top + ptMargin '// ��[
    basePos(1) = targetShape.TopLeftCell.Left + ptMargin '// ���[
    basePos(2) = targetShape.BottomRightCell.Top + targetShape.BottomRightCell.Height - ptMargin '// ���[
    basePos(3) = targetShape.BottomRightCell.Left + targetShape.BottomRightCell.Width - ptMargin '// �E�[
    
    If targetShape.Type <> msoLine Then   '// �����V�F�C�v�ȊO��ΏۂƂ���
        '// �㉺�[�ݒ�
        If Not ckbDetail.Value Or (cmbHeight.Value = 0) Or (cmbHeight.Value = 1) Then
            targetShape.Top = basePos(0)
            If cmbHeight.Value = 0 Then
                targetShape.Height = basePos(2) - basePos(0)
            End If
        ElseIf cmbHeight.Value = 2 Then
            targetShape.Top = basePos(2) - targetShape.Height
        End If
        
        '// ���E�[�ݒ�
        If Not ckbDetail.Value Or (cmbWidth.Value = 0) Or (cmbWidth.Value = 1) Then
            targetShape.Left = basePos(1)
            If cmbWidth.Value = 0 Then
                targetShape.Width = basePos(3) - basePos(1)
            End If
        ElseIf cmbWidth.Value = 2 Then
            targetShape.Left = basePos(3) - targetShape.Width
        End If
    End If
End Sub

'// ////////////////////////////////////////////////////////////////////////////
'// END
'// ////////////////////////////////////////////////////////////////////////////
