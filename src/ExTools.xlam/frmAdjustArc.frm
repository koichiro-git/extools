VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAdjustArc 
   Caption         =   "�~�ʂ̒���"
   ClientHeight    =   2565
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3030
   OleObjectBlob   =   "frmAdjustArc.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "frmAdjustArc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'// ////////////////////////////////////////////////////////////////////////////
'// �v���W�F�N�g   : �g���c�[��
'// �^�C�g��       : �~�ʂ̒����t�H�[��
'// ���W���[��     : frmAdjustArc
'// ����           : �~�ʃI�u�W�F�N�g�̊J�n�ʒu�A�I���ʒu���p�x�Ŏw�肷��
'//                : �ΏۂƂ���V�F�C�v�� Pie, BlockArc, CircularArrow, msoShapeArc
'// ////////////////////////////////////////////////////////////////////////////
'// Copyright (c) by Koichiro.
'// ////////////////////////////////////////////////////////////////////////////
Option Explicit
Option Base 0

'// ////////////////////////////////////////////////////////////////////////////
'// �R���p�C���X�C�b�`�i"EXCEL" / "POWERPOINT"�j
'#Const OFFICE_APP = "EXCEL"

Private Const ANGLE_ADJUST      As Integer = -90    '// �p�x�v�Z�̊J�n�ʒu�␳�l

Private angleStart              As Integer          '// �J�n�p�x
Private angleEnd                As Integer          '// �I���p�x


'// //////////////////////////////////////////////////////////////////
'// �C�x���g�F �t�H�[�� ��������
Private Sub UserForm_Initialize()
    '// �L���v�V�����ݒ�
    Me.Caption = LBL_ARC_FORM
    lblStart.Caption = LBL_ARC_START
    lblEnd.Caption = LBL_ARC_END
    cmdResetRotation.Caption = LBL_ARC_RESET_ROT
    cmdClose.Caption = LBL_COM_CLOSE
End Sub


'// //////////////////////////////////////////////////////////////////
'// �C�x���g�F �t�H�[�� �A�N�e�B�x�[�g��
Private Sub UserForm_Activate()
On Error GoTo ErrorHandler
    Dim shp     As Shape
    
    '// ���O���� //////////
    '// ���O�`�F�b�N�i�A�N�e�B�u�V�[�g�ی�A�I���^�C�v���V�F�C�v�j
    If Not gfPreCheck(protectCont:=True, selType:=TYPE_SHAPE) Then
        Me.Hide
        Exit Sub
    End If
    
    For Each shp In ActiveWindow.Selection.ShapeRange
        Select Case shp.AutoShapeType
            Case msoShapePie, msoShapeBlockArc, msoShapeArc
                angleStart = Int(shp.Adjustments.Item(1)) - ANGLE_ADJUST
                angleEnd = Int(shp.Adjustments.Item(2)) - ANGLE_ADJUST
                Exit For
            Case msoShapeCircularArrow
                angleStart = Int(shp.Adjustments.Item(4)) - ANGLE_ADJUST
                angleEnd = Int(shp.Adjustments.Item(3)) - ANGLE_ADJUST + shp.Adjustments.Item(2)
                Exit For
        End Select
    Next
    txtStart.Value = angleStart
    txtEnd.Value = angleEnd
    
    Exit Sub

ErrorHandler:
    Call gsShowErrorMsgDlg("frmAdjustArc.UserForm_Activate", Err, Nothing)
End Sub


'// //////////////////////////////////////////////////////////////////
'// �C�x���g�F ����{�^�� �N���b�N��
Private Sub cmdClose_Click()
    Call Me.Hide
End Sub


'// //////////////////////////////////////////////////////////////////
'// �C�x���g�F �X�����Z�b�g�{�^�� �N���b�N��
Private Sub cmdResetRotation_Click()
    Dim shp As Shape
    
    For Each shp In ActiveWindow.Selection.ShapeRange
        Select Case shp.AutoShapeType
            Case msoShapePie, msoShapeBlockArc, msoShapeCircularArrow, msoShapeArc
                shp.Rotation = 0
        End Select
    Next
End Sub


'// //////////////////////////////////////////////////////////////////
'// �C�x���g�F�X�s���{�^��(�J�n�p�x)
'// SpinUp
Private Sub spnStart_SpinUp()
    angleStart = Int(angleStart / 15) * 15 + 15
    txtStart.Value = angleStart
    
    Call adjustArc
End Sub
'// SpinDown
Private Sub spnStart_SpinDown()
    If angleStart = Int(angleStart / 15) * 15 Then
        angleStart = Int(angleStart / 15) * 15 - 15
    Else
        angleStart = Int(angleStart / 15) * 15
    End If
    
    txtStart.Value = angleStart
    Call adjustArc
End Sub


'// //////////////////////////////////////////////////////////////////
'// �C�x���g�F�X�s���{�^��(�I���p�x)
'// SpinUp
Private Sub spnEnd_SpinUp()
    angleEnd = Int(angleEnd / 15) * 15 + 15
    txtEnd.Value = angleEnd
    
    Call adjustArc
End Sub
'// SpinDown
Private Sub spnEnd_SpinDown()
    If angleEnd = Int(angleEnd / 15) * 15 Then
        angleEnd = Int(angleEnd / 15) * 15 - 15
    Else
        angleEnd = Int(angleEnd / 15) * 15
    End If
    
    txtEnd.Value = angleEnd
    Call adjustArc
End Sub


'// //////////////////////////////////////////////////////////////////
'// �C�x���g�F �J�n�p�x�e�L�X�g�{�b�N�X AfterUpdate��
Private Sub txtStart_AfterUpdate()
On Error GoTo ErrorHandler
    If IsNumeric(txtStart.Value) Then
        txtStart.Value = Int(txtStart.Value)
        angleStart = Int(txtStart.Value)
        Call adjustArc
    Else
        Call MsgBox(MSG_INVALID_NUM, vbOKOnly, APP_TITLE)
        txtStart.Value = ActiveWindow.Selection.ShapeRange(1).Adjustments.Item(1) - ANGLE_ADJUST
    End If
    Exit Sub
    
ErrorHandler:
    Call gsShowErrorMsgDlg("frmAdjustArc.txtStart_AfterUpdate", Err, Nothing)
End Sub


'// //////////////////////////////////////////////////////////////////
'// �C�x���g�F �I���p�x�e�L�X�g�{�b�N�X AfterUpdate��
Private Sub txtEnd_AfterUpdate()
On Error GoTo ErrorHandler
    If IsNumeric(txtEnd.Value) Then
        txtEnd.Value = Int(txtEnd.Value)
        angleEnd = Int(txtEnd.Value)
        Call adjustArc
    Else
        Call MsgBox(MSG_INVALID_NUM, vbOKOnly, APP_TITLE)
        txtEnd.Value = ActiveWindow.Selection.ShapeRange(1).Adjustments.Item(2) - ANGLE_ADJUST
    End If
    Exit Sub
    
ErrorHandler:
    Call gsShowErrorMsgDlg("frmAdjustArc.txtEnd_AfterUpdate", Err, Nothing)
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// ���\�b�h�F   �~�ʊp�x����
'// �����F       �I������Ă���~�ʃI�u�W�F�N�g�̊J�n�p�x�E�I���p�x��ݒ肷��
'// ////////////////////////////////////////////////////////////////////////////
Private Sub adjustArc()
    Dim shp     As Shape
    
    For Each shp In ActiveWindow.Selection.ShapeRange
        Select Case shp.AutoShapeType
            Case msoShapePie, msoShapeBlockArc, msoShapeArc
                shp.Adjustments.Item(1) = angleStart + ANGLE_ADJUST
                shp.Adjustments.Item(2) = angleEnd + ANGLE_ADJUST
            Case msoShapeCircularArrow
                shp.Adjustments.Item(4) = angleStart + ANGLE_ADJUST
                shp.Adjustments.Item(3) = angleEnd + ANGLE_ADJUST - shp.Adjustments.Item(2)
        End Select
    Next
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// END
'// ////////////////////////////////////////////////////////////////////////////
