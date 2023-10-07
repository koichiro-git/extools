Attribute VB_Name = "mdlConvVal"
'// ////////////////////////////////////////////////////////////////////////////
'// �v���W�F�N�g   : �g���c�[��
'// �^�C�g��       : ������̕ϊ�
'// ���W���[��     : mdlConvVal
'// ����           : �Z��������̕ϊ��A�g�����@�\
'// ////////////////////////////////////////////////////////////////////////////
'// Copyright (c) by Koichiro.
'// ////////////////////////////////////////////////////////////////////////////
Option Explicit
Option Base 0


'// ////////////////////////////////////////////////////////////////////////////
'// ���\�b�h�F   ���{���{�^���R�[���o�b�N�Ǘ�(�t�H�[���Ȃ�)
'// �����F       ���{������̃R�[���o�b�N�������ǂ�
'//              �����ꂽ�R���g���[����ID����ɏ������Ăяo���B
'// �����F       control �ΏۃR���g���[��
'// ////////////////////////////////////////////////////////////////////////////
Public Sub ribbonCallback_ConvVal(control As IRibbonControl)
    Select Case control.ID
        Case "chrUpper"                     '// �啶��
            Call psConvValue(MENU_CAPITAL)
        Case "chrLower"                     '// ������
            Call psConvValue(MENU_SMALL)
        Case "chrInitCap"                   '// �擪�啶��
            Call psConvValue(MENU_PROPER)
        Case "chrZen"                       '// �S�p
            Call psConvValue(MENU_ZEN)
        Case "chrHan"                       '// ���p
            Call psConvValue(MENU_HAN)
        Case "TrimVal"                      '// �g����
            Call psConvValue(MENU_TRIM)
    End Select
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// ���\�b�h�F   ������̕ϊ�
'// �����F       �I��͈͂̒l��ϊ�����
'// ////////////////////////////////////////////////////////////////////////////
Private Sub psConvValue(funcFlag As String)
On Error GoTo ErrorHandler
    Dim tCell     As Range    '// �ϊ��ΏۃZ��
    Dim statGauge As cStatusGauge
    
    '// ���O�`�F�b�N�i�A�N�e�B�u�V�[�g�ی�A�I���^�C�v���Z���j
    If Not gfPreCheck(protectCont:=True, selType:=TYPE_RANGE) Then
        Exit Sub
    End If
    
    Call gsSuppressAppEvents
    
    If Selection.Count > 1 Then
        For Each tCell In Selection.SpecialCells(xlCellTypeConstants, xlNumbers + xlTextValues)
            Call psConvValue_sub(tCell, funcFlag)
            
            '// �L�[����
            If GetAsyncKeyState(27) <> 0 Then
                Application.StatusBar = False
                Exit For
            End If
        Next
    Else
        Call psConvValue_sub(ActiveCell, funcFlag)
    End If
    
    Call gsResumeAppEvents
    Exit Sub
ErrorHandler:
    Call gsResumeAppEvents
    If Err.Number = 1004 Then  '// �͈͑I�����������Ȃ��ꍇ
        Call MsgBox(MSG_INVALID_RANGE, vbOKOnly, APP_TITLE)
    Else
        Call gsShowErrorMsgDlg("psConvValue", Err)
    End If
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// ���\�b�h�F   ������̕ϊ� �T�u���[�`��
'// �����F       �����̒l��ϊ�����
'// ////////////////////////////////////////////////////////////////////////////
Private Sub psConvValue_sub(tCell As Range, funcFlag As String)
    Select Case funcFlag
        Case MENU_CAPITAL
            tCell.Value = UCase(tCell.Value)
        Case MENU_SMALL
            tCell.Value = LCase(tCell.Value)
        Case MENU_PROPER
            tCell.Value = StrConv(tCell.Value, vbProperCase)
        Case MENU_ZEN
            tCell.Value = StrConv(tCell.Value, vbWide)
        Case MENU_HAN
            tCell.Value = StrConv(StrConv(tCell.Value, vbKatakana), vbNarrow)
        Case MENU_TRIM
            tCell.Value = Trim$(tCell.Value)
            If Len(tCell.Value) = 0 Then
                tCell.Value = Empty
            End If
    End Select
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// END
'// ////////////////////////////////////////////////////////////////////////////
