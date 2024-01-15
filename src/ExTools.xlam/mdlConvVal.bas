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

Private Const MENU_CHANGE_CHAR                   As String = "Change Case"
Private Const MENU_CAPITAL                       As String = "Uppercase"
Private Const MENU_SMALL                         As String = "Lowercase"
Private Const MENU_PROPER                        As String = "Capital the First Letter in the Word"
Private Const MENU_ZEN                           As String = "Wide Letter"
Private Const MENU_HAN                           As String = "Narrow Letter"
Private Const MENU_TRIM                          As String = "Trim Values"

'// ////////////////////////////////////////////////////////////////////////////
'// �� �錾
'// �^�ϊ����[�h
Public Enum udConvMode
  cUpper = 1
  cSmall = 2
  cProper = 3
  cZenkaku = 4
  cHankaku = 5
  cTrim = 6
End Enum


'// ////////////////////////////////////////////////////////////////////////////
'// ���\�b�h�F   ���{���{�^���R�[���o�b�N�Ǘ�(�t�H�[���Ȃ�)
'// �����F       ���{������̃R�[���o�b�N�������ǂ�
'//              �����ꂽ�R���g���[����ID����ɏ������Ăяo���B
'// �����F       control �ΏۃR���g���[��
'// ////////////////////////////////////////////////////////////////////////////
Public Sub ribbonCallback_ConvVal(control As IRibbonControl)
    Select Case control.ID
        Case "chrUpper"                     '// �啶��
            Call psConvValue(cUpper)
        Case "chrLower"                     '// ������
            Call psConvValue(cSmall)
        Case "chrInitCap"                   '// �擪�啶��
            Call psConvValue(cProper)
        Case "chrZen"                       '// �S�p
            Call psConvValue(cZenkaku)
        Case "chrHan"                       '// ���p
            Call psConvValue(cHankaku)
        Case "TrimVal"                      '// �g����
            Call psConvValue(cTrim)
    End Select
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// ���\�b�h�F   ������̕ϊ�
'// �����F       �I��͈͂̒l��ϊ�����
'// ////////////////////////////////////////////////////////////////////////////
Private Sub psConvValue(funcFlag As udConvMode)
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
        Call gsShowErrorMsgDlg("psConvValue", Err, Nothing)
    End If
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// ���\�b�h�F   ������̕ϊ� �T�u���[�`��
'// �����F       �����̒l��ϊ�����
'// ////////////////////////////////////////////////////////////////////////////
Private Sub psConvValue_sub(tCell As Range, funcFlag As udConvMode)
    Select Case funcFlag
        Case cUpper
            tCell.Value = UCase(tCell.Value)
        Case cSmall
            tCell.Value = LCase(tCell.Value)
        Case cProper
            tCell.Value = StrConv(tCell.Value, vbProperCase)
        Case cZenkaku
            tCell.Value = StrConv(tCell.Value, vbWide)
        Case cHankaku
            tCell.Value = StrConv(StrConv(tCell.Value, vbKatakana), vbNarrow)
        Case cTrim
            tCell.Value = Trim$(tCell.Value)
            If Len(tCell.Value) = 0 Then
                tCell.Value = Empty
            End If
    End Select
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// END
'// ////////////////////////////////////////////////////////////////////////////
