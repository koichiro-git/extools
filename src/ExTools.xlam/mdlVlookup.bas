Attribute VB_Name = "mdlVlookup"
'// ////////////////////////////////////////////////////////////////////////////
'// �v���W�F�N�g   : �g���c�[��
'// �^�C�g��       : Vlookup�\��t���@�\
'// ���W���[��     : mdlVlookup
'// ����           : VLOOKUP�̃}�X�^��`�̋L���A�����VLOOKUP�֐��̓\��t�����s��
'// ////////////////////////////////////////////////////////////////////////////
'// Copyright (c) by Koichiro.
'// ////////////////////////////////////////////////////////////////////////////
Option Explicit
Option Base 0

'// ////////////////////////////////////////////////////////////////////////////
'// �v���C�x�[�g�ϐ�
Private pVLookupMaster                  As String               '// VLookUp�R�s�[�@�\�Ń}�X�^�\�͈͂��i�[����
Private pVLookupMasterIndex             As String               '// VLookUp�R�s�[�@�\�Ń}�X�^�\�͈͂̕\���C���f�N�X���i�[����v


'// ////////////////////////////////////////////////////////////////////////////
'// ���\�b�h�F   ���{���{�^���R�[���o�b�N�Ǘ�(�t�H�[���Ȃ�)
'// �����F       ���{������̃R�[���o�b�N�������ǂ�
'//              �����ꂽ�R���g���[����ID����ɏ������Ăяo���B
'// �����F       control �ΏۃR���g���[��
'// ////////////////////////////////////////////////////////////////////////////
Public Sub ribbonCallback_Vlookup(control As IRibbonControl)
    Select Case control.ID
        Case "VLookupCopy"                  '// VLookup
            Call psVLookupCopy
        Case "VLookupPaste"
            Call psVLookupPaste
    End Select
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// ���\�b�h�F   VLookup�̃}�X�^�̈�Ƃ��ăR�s�[
'// �����F       �I��̈��\�������������ϐ��Ɋi�[����
'// ////////////////////////////////////////////////////////////////////////////
Private Sub psVLookupCopy()
    '// ���O�`�F�b�N
    If gfPreCheck(selType:=TYPE_RANGE, selAreas:=1) = False Then
        Exit Sub
    End If
    
    '// 1��݂̂̑I���̓G���[
    If Selection.Columns.Count = 1 Then
        Call MsgBox(MSG_VLOOKUP_MASTER_2COLS, vbOKOnly, APP_TITLE)
        Exit Sub
    End If
    
    pVLookupMaster = Selection.Worksheet.Name & "!" & Selection.Address(True, True)   '// ��ƍs���ΎQ��
    pVLookupMasterIndex = Selection.Columns.Count
End Sub
            
            
'// ////////////////////////////////////////////////////////////////////////////
'// ���\�b�h�F   VLookup�֐��𒣂�t��
'// �����F       VLookupCopy�Ŋi�[���ꂽ�I��̈�𒣂�t���ʒu�̃Z���ɏo�͂���
'// ////////////////////////////////////////////////////////////////////////////
Private Sub psVLookupPaste()
On Error GoTo ErrorHandler
    Dim searchColIdx    As Long     '// VLookup�֐��́u�����v�ɂ�����Z���̗�
    Dim targetColIdx    As Long     '// Vlookpu�֐����o�͂���Z���̗�
    Dim bffRange        As String   '// �I��͈͂̃A�h���X�������ێ�
    
    '// ���O�`�F�b�N
    If gfPreCheck(protectCont:=True, selType:=TYPE_RANGE, selAreas:=1) = False Then
        Exit Sub
    End If
    
    '// �}�X�^�񂪑I������Ă��Ȃ��ꍇ�̓G���[
    If pVLookupMaster = BLANK Then
        Call MsgBox(MSG_VLOOKUP_NO_MASTER, vbOKOnly, APP_TITLE)
        Exit Sub
    End If
    
    '// 1��݂̂̑I���̓G���[
    If Selection.Columns.Count = 1 Then
        Call MsgBox(MSG_VLOOKUP_SET_2COLS, vbOKOnly, APP_TITLE)
        Exit Sub
    End If
    
    '// ����V�[�g���ł̓\��t���̏ꍇ�̓}�X�^�\�͈̔͂Ƃ̏d�����`�F�b�N
    If Selection.Worksheet.Name = Range(pVLookupMaster).Worksheet.Name Then
        If Not Application.Intersect(Selection, Range(Range(pVLookupMaster).Address)) Is Nothing Then
            Call MsgBox(MSG_VLOOKUP_SEL_DUPLICATED, vbOKOnly, APP_TITLE)
            Exit Sub
        End If
    End If
    
    '// �I��͈͂̂����A�J�����g�Z����VLOOKUP�́u�����v��ɊY��
    searchColIdx = ActiveCell.Column
    '// ���ۂ�VLOOKUP�֐��𖄂ߍ��ރZ���́A�I��͈͂̍Ō㑤�B
    targetColIdx = IIf(Selection.Column = ActiveCell.Column, Selection.Column + Selection.Columns.Count - 1, Selection.Column)
    
    '// �ŏ��̃Z�����o��
    Cells(Selection.Row, targetColIdx).Value = "=VLOOKUP(" & ActiveCell.Address(False, False) & "," & pVLookupMaster & "," & Str(pVLookupMasterIndex) & ",FALSE)"
    '// �������ɐ����̂݃R�s�[
    bffRange = Selection.Address(False, False)
    Call Cells(Selection.Row, targetColIdx).Copy
    Call Range(Cells(Selection.Row, targetColIdx), Cells(Selection.Row + Selection.Rows.Count - 1, targetColIdx)).PasteSpecial(Paste:=xlPasteFormulas, Operation:=xlNone, SkipBlanks:=False, Transpose:=False)
    
    '// �㏈��
    Application.CutCopyMode = False
    Exit Sub
    
ErrorHandler:
    Call gsShowErrorMsgDlg("psVLookupPaste", Err)
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// END
'// ////////////////////////////////////////////////////////////////////////////

