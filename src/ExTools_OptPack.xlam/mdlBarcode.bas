Attribute VB_Name = "mdlBarcode"
'// ////////////////////////////////////////////////////////////////////////////
'// �v���W�F�N�g   : �g���c�[�� �ǉ��p�b�N
'// �^�C�g��       : QR�R�[�h�\��
'// ���W���[��     : mdlBarcode
'// ����           : �Z�����e����QR�R�[�h�摜�𐶐�����B
'//                  Access�p�����^�C����Excel�ł͓���ۏ؂���Ă��Ȃ����߁A�t���[��Web�T�[�r�X���g�p
'// ////////////////////////////////////////////////////////////////////////////
'// Copyright (c) by Koichiro.
'// ////////////////////////////////////////////////////////////////////////////
Option Explicit
Option Base 0


'// ////////////////////////////////////////////////////////////////////////////
'// ���\�b�h�F   ���{���{�^���R�[���o�b�N�Ǘ�
'// �����F       ���{������̃R�[���o�b�N�������ǂ�
'//              �����ꂽ�R���g���[����ID����ɏ������Ăяo���B
'// �����F       control �ΏۃR���g���[��
'// ////////////////////////////////////////////////////////////////////////////
Public Sub ribbonCallback_BarCode(control As IRibbonControl)
    Select Case control.ID
        Case "QRCode"                       '// QR�R�[�h
            Call psDrawBarCode
    End Select
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// ���\�b�h�F   �o�[�R�[�h�`��
'// �����F       �I�����ꂽ�Z���̒l�����Ƀo�[�R�[�h��`�悷��
'// ////////////////////////////////////////////////////////////////////////////
Private Sub psDrawBarCode()
    Dim tCell     As Range    '// �ϊ��ΏۃZ��
    
    '// ���O�`�F�b�N�i�A�N�e�B�u�V�[�g�ی�A�I���^�C�v���Z���j
    If Not gfPreCheck(protectCont:=True, selType:=TYPE_RANGE) Then
        Exit Sub
    End If
    
    Call gsSuppressAppEvents
    
    If Selection.Count > 1 Then
        For Each tCell In Selection.SpecialCells(xlCellTypeConstants, xlNumbers + xlTextValues)
            Call psDrawBarCode_sub(tCell)
        Next
    Else
        Call psDrawBarCode_sub(ActiveCell)
    End If
    
    Call gsResumeAppEvents
    Exit Sub
    
ErrorHandler:
    Call gsResumeAppEvents
    Call gsShowErrorMsgDlg_VBA("mdlBarcode.psDrawBarCode", Err)
End Sub

    
'// ////////////////////////////////////////////////////////////////////////////
'// ���\�b�h�F   �o�[�R�[�h�`��i�T�u�j
'// �����F       �t���[��QR�R�[�h�`��T�[�r�X�ɃA�N�Z�X���A�����̃Z���ɃV�F�C�v�Ƃ��Ē���t����
'// �����F       tCell: �V�F�C�v�\��t���ʒu
'// ////////////////////////////////////////////////////////////////////////////
Private Sub psDrawBarCode_sub(tCell As Range)
    Dim obj         As Shape    '// �摜�\�t�p�V�F�C�v
    
    Set obj = ActiveSheet.Shapes.AddPicture("http://api.qrserver.com/v1/create-qr-code/?data=" & tCell.Text & _
                                            "!&size=300x300", False, True, tCell.Left, tCell.Top, tCell.Width, tCell.Height)
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// END
'// ////////////////////////////////////////////////////////////////////////////
