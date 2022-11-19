Attribute VB_Name = "mdlConvJapaneseNumber"
'// ////////////////////////////////////////////////////////////////////////////
'// �v���W�F�N�g   : �g���c�[�� �ǉ��p�b�N
'// �^�C�g��       : ���{��\�L���l�ϊ�
'// ���W���[��     : mdlConvJapaneseNumber
'// ����           : �V�X�e���̋��ʊ֐��A�N�����̐ݒ�Ȃǂ��Ǘ�
'//                  �����d�b�ԍ��̂ݑΉ��B���ԍ��i+81���j�͑Ή����Ă��Ȃ�
'// ////////////////////////////////////////////////////////////////////////////
'// Copyright (c) by Koichiro.
'// ////////////////////////////////////////////////////////////////////////////
Option Explicit
Option Base 0

Public Sub test()
    Dim Text As String
    Call pfConvJapaneseToNumber
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// ���\�b�h�F   ���{���{�^���R�[���o�b�N�Ǘ�
'// �����F       ���{������̃R�[���o�b�N�������ǂ�
'//              �����ꂽ�R���g���[����ID����ɏ������Ăяo���B
'// �����F       control �ΏۃR���g���[��
'// ////////////////////////////////////////////////////////////////////////////
Public Sub ribbonCallback_JapaneseNum(control As IRibbonControl)
    Call pfConvJapaneseToNumber
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// ���\�b�h�F   ���{��\�L���l�ϊ� ��֐�
'// �����F       �����̓��{��\�L���l�𐔒l�ɕϊ�����i5��4��3�S �� 54300�j
'//              �����͔�Ή�
'// ////////////////////////////////////////////////////////////////////////////
Private Sub pfConvJapaneseToNumber()
On Error GoTo ErrorHandler
    Dim tCell       As Range    '// �ϊ��ΏۃZ��
    Dim bff         As String   '// �ϊ��㕶����i�[�o�b�t�@
    
    '// ���O�`�F�b�N�i�A�N�e�B�u�V�[�g�ی�A�I���^�C�v���Z���j
    If Not gfPreCheck(protectCont:=True, selType:=TYPE_RANGE) Then
        Exit Sub
    End If
    
    Call gsSuppressAppEvents
    
    If Selection.Count > 1 Then
        For Each tCell In Selection.SpecialCells(xlCellTypeConstants, xlNumbers + xlTextValues)  '//SELECTION����̏ꍇ�̓G���[�n���h���ŃL���b�`
            bff = pfConvJapaneseToNumber_sub(tCell.Text)
            If bff <> BLANK Then    '// �ϊ����W�b�N����u�����N���߂��ꂽ�ꍇ�͖���
                tCell.Value = bff
            End If
        Next
    Else
        bff = pfConvJapaneseToNumber_sub(ActiveCell.Text)
        If bff <> BLANK Then        '// �ϊ����W�b�N����u�����N���߂��ꂽ�ꍇ�͖���
            ActiveCell.Value = bff
        End If
    End If
    
    Call gsResumeAppEvents
    Exit Sub
ErrorHandler:
    Call gsResumeAppEvents
    If Err.Number = 1004 Then  '// �͈͑I�����������Ȃ��ꍇ
        Call MsgBox(MSG_INVALID_RANGE, vbOKOnly, APP_TITLE)
    Else
        Call gsShowErrorMsgDlg_VBA("mdlConvJapaneseNumber.pfConvJapaneseToNumber", Err)
    End If
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// ���\�b�h�F   ���{��\�L���l�ϊ�
'// �����F       �����̓��{��\�L���l�𐔒l�ɕϊ�����i5��4��3�S �� 54300�j
'//              psFormatPhoneNumbers ����Ăяo����������
'// �����F       ���{��\�L���l
'// ////////////////////////////////////////////////////////////////////////////
Private Function pfConvJapaneseToNumber_sub(targetStr As String) As Double
On Error GoTo ErrorHandler
'    Dim targetStr   As String
    Dim i           As Integer  '// �����𒊏o����ۂ̃C���f�N�X
    Dim c           As String   '// ���o���ꂽ�ꕶ�����i�[����o�b�t�@
    Dim dig         As Double   '// ���݂̈�
    Dim lastKanji   As String   '// ���������ꂽ�ꍇ�A���̍Ō�̊�����ێ��i���A���A���̂����ꂩ�j
    Dim Result      As Double   '// �ϊ��̌o�߂�ێ�����

'    targetStr = "123���S��5�S��3�疜5��"
'    targetStr = "2��3�疜�~"
'    targetStr = "230000000"
'    targetStr = "2"
'    targetStr = "-2�S"
'    targetStr = "12��400��325��5��4"
'    targetStr = "5��4��"
'    targetStr = "��12�R��4��00��3��2�S5��5��4�S"
    dig = 1
    i = Len(targetStr)
    Do While i > 0
        c = Mid(targetStr, i, 1)
        
        If IsNumeric(c) Then
            Result = Result + dig * Int(c)
            dig = dig * 10
        Else
            Select Case c
                Case "�S"
                    Select Case lastKanji
                        Case ""
                            dig = 100
                        Case "��"
                            dig = 1000000
                        Case "��"
                            dig = 10000000000#
                        Case "��"
                            dig = 100000000000000#
                    End Select
                Case "��"
                    Select Case lastKanji
                        Case ""
                            dig = 1000
                        Case "��"
                            dig = 10000000
                        Case "��"
                            dig = 100000000000#
                        Case "��"
                            '// 100���ȏ�̓G���[
                    End Select
                Case "��"
                    dig = 10000
                    lastKanji = "��"
                Case "��"
                    dig = 100000000
                    lastKanji = "��"
                Case "��"
                    dig = 1000000000000#
                    lastKanji = "��"
                '// ����ȊO�̕����͖���
            End Select
        End If
        
        i = i - 1
    Loop
    '// �}�C�i�X�Ή�
    If InStr(1, "-����", Left(targetStr, 1)) > 0 Then
        Result = Result * -1
    End If
    pfConvJapaneseToNumber_sub = Result
    Exit Function
    
ErrorHandler:
    Call gsResumeAppEvents
    Call gsShowErrorMsgDlg_VBA("mdlConvJapaneseNumber.pfConvJapaneseToNumber_sub", Err)
End Function


'// ////////////////////////////////////////////////////////////////////////////
'// END
'// ////////////////////////////////////////////////////////////////////////////

