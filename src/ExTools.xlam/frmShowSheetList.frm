VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmShowSheetList 
   Caption         =   "�V�[�g�ꗗ�o��"
   ClientHeight    =   3375
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4305
   OleObjectBlob   =   "frmShowSheetList.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "frmShowSheetList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'// ////////////////////////////////////////////////////////////////////////////
'// �v���W�F�N�g   : �g���c�[��
'// �^�C�g��       : �V�[�g�ꗗ�o�̓t�H�[��
'// ���W���[��     : frmShowSheetList
'// ����           : �V�[�g�ꗗ���o�͂���
'// ////////////////////////////////////////////////////////////////////////////
'// Copyright (c) by Koichiro.
'// ////////////////////////////////////////////////////////////////////////////
Option Explicit
Option Base 0


'// ////////////////////////////////////////////////////////////////////////////
'// �v���C�x�[�g�萔
Private Const pDEF_ROWS  As Integer = 5
Private Const pDEF_COLS  As Integer = 5
Private Const pMAX_ROWS  As Integer = 10  '// �s���R���{�{�b�N�X�̍ő�l
Private Const pMAX_COLS  As Integer = 60  '// �񐔃R���{�{�b�N�X�̍ő�l


'// //////////////////////////////////////////////////////////////////
'// �C�x���g�F �t�H�[�� �A�N�e�B�u��
Private Sub UserForm_Activate()
    '// �u�b�N���J����Ă��Ȃ��ꍇ�͏I��
    If Workbooks.Count = 0 Then
        Call MsgBox(MSG_NO_BOOK, vbOKOnly, APP_TITLE)
        Call Me.Hide
        Exit Sub
    End If
    
    If ActiveWorkbook.MultiUserEditing Or (cmbOutput.Value = "0") Then
        ckbHyperLink.Value = False
        ckbHyperLink.Enabled = False
    Else
        ckbHyperLink.Enabled = True
    End If
End Sub


'// //////////////////////////////////////////////////////////////////
'// �C�x���g�F �t�H�[�� ��������
Private Sub UserForm_Initialize()
    Dim idx   As Integer
    
    Call gsSetCombo(cmbOutput, CMB_SSL_OUTPUT, 0)
    
    '// �s�E��R���{�̐ݒ�
    With cmbRows
        Call .Clear
        For idx = 0 To pMAX_ROWS
            Call .AddItem(CStr(idx))
            .List(idx, 1) = CStr(idx)
        Next
        .ListIndex = pDEF_ROWS
    End With
    
    With cmbCols
        Call .Clear
        For idx = 0 To pMAX_COLS
            Call .AddItem(CStr(idx))
            .List(idx, 1) = CStr(idx)
        Next
        .ListIndex = pDEF_COLS
    End With
    
    '// �L���v�V�����ݒ�
    frmShowSheetList.Caption = LBL_SSL_FORM
    cmdExecute.Caption = LBL_COM_EXEC
    cmdClose.Caption = LBL_COM_CLOSE
    fraOption.Caption = LBL_SSL_OPTIONS
    ckbHyperLink.Caption = LBL_COM_HYPERLINK
    lblTarget.Caption = LBL_SSL_TARGET
    lblRows.Caption = LBL_SSL_ROWS
    lblCols.Caption = LBL_SSL_COLS
End Sub


'// //////////////////////////////////////////////////////////////////
'// �C�x���g�F ���s�{�^�� �N���b�N��
Private Sub cmdExecute_Click()
    '// �u�b�N���ی삳��Ă���ꍇ�ŁA����u�b�N�Ɍ��ʃV�[�g��ǉ����悤�Ƃ����ꍇ�̓G���[
    If ActiveWorkbook.ProtectStructure And (cmbOutput.Value <> "0") Then
        Call MsgBox(MSG_BOOK_PROTECTED, vbOKOnly, APP_TITLE)
    Else
        '// �������{
        Call psShowSheetList(ActiveWorkbook, cmbRows.Value, cmbCols.Value)
    End If
    Call Me.Hide
End Sub


'// //////////////////////////////////////////////////////////////////
'// �C�x���g�F ����{�^�� �N���b�N��
Private Sub cmdClose_Click()
    Call Me.Hide
End Sub


'// //////////////////////////////////////////////////////////////////
'// �C�x���g�F �o�͐�R���{ �ύX��
Private Sub cmbOutput_Change()
    '// �u�b�N���J����Ă��Ȃ��ꍇ�͂Ȃɂ������I��
    If Workbooks.Count = 0 Then
        Exit Sub
    End If

    If cmbOutput.Value = "0" Then
        ckbHyperLink.Value = False
        ckbHyperLink.Enabled = False
    ElseIf Not ActiveWorkbook.MultiUserEditing Then
        ckbHyperLink.Enabled = True
    End If
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// ���\�b�h�F   �V�[�g�ꗗ�o��
'// �����F       �V�[�g�ꗗ���o�͂���
'//              ���s�{�^���N���b�N�C�x���g����Ăяo�����B
'// �����F       wkBook: �Ώۃu�b�N
'//              maxRow: �o�͍s��
'//              maxCol: �o�͗�
'// ////////////////////////////////////////////////////////////////////////////
Private Sub psShowSheetList(wkBook As Workbook, maxRow As Integer, maxCol As Integer)
    Dim resultSheet As Worksheet  '// ���ʏo�͐�̃V�[�g
    Dim sheetObj    As Object     '// worksheet �܂��� chart �I�u�W�F�N�g���i�[
    Dim idx         As Integer    '// ���ʃV�[�g�̃J�����ʒu�␳
    Dim idxRow      As Integer
    Dim idxCol      As Integer
    Dim statGauge   As cStatusGauge
    
    '// �u�b�N���J����Ă��Ȃ��ꍇ�͏I��
    If Workbooks.Count = 0 Then
        Call MsgBox(MSG_NO_BOOK, vbOKOnly, APP_TITLE)
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    Set statGauge = New cStatusGauge
    statGauge.MaxVal = wkBook.Sheets.Count * maxRow
  
  '// �o�͐�̐ݒ�
    Select Case cmbOutput.Value
        Case "0"
            Call Workbooks.Add
            Set resultSheet = ActiveWorkbook.ActiveSheet
        Case "1"
            Set resultSheet = ActiveWorkbook.Worksheets.Add(Before:=ActiveWorkbook.Worksheets(1))
        Case "2"
            Set resultSheet = ActiveWorkbook.Worksheets.Add(After:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count))
    End Select
  
    '// B��̓V�[�g���̂̂��߁A�����𕶎���i@�j�ɐݒ�
    resultSheet.Columns("B").NumberFormat = "@"
    
    '// �w�b�_�̐ݒ�(�Z���C���f�N�X�̕\���͋��ʊ֐����g�p���Ȃ��B�ŏ���2�񕪁i"�V�[�g�ԍ�;�V�[�g����"�j�݂̂�HDR_SSL�Őݒ�)
    Call gsDrawResultHeader(resultSheet, HDR_SSL, 1)
    '// 3��ڈȍ~�̃w�b�_�́uA1,B1,C1...�v���R���{�{�b�N�X�ł̎w��񕪐ݒ�
    idx = 3
    For idxRow = 1 To maxRow
        For idxCol = 1 To maxCol
            resultSheet.Cells(1, idx).Value = gfGetColIndexString(idxCol) & CStr(idxRow)
            idx = idx + 1
        Next
    Next
  
    '// �ꗗ�i�f�[�^���j�̏o��
    For Each sheetObj In wkBook.Sheets
        resultSheet.Cells(sheetObj.Index + 1, 1).Value = sheetObj.Index
        resultSheet.Cells(sheetObj.Index + 1, 2).Value = sheetObj.Name
        
        If sheetObj.Type = xlWorksheet Then '// ���[�N�V�[�g�̂݁A���e�̕\���ƃ����N�̐ݒ�
            '// �����N�̐ݒ�
            If ckbHyperLink.Value And (sheetObj.Visible = xlSheetVisible) Then
                Call Cells(resultSheet.Index + 1, 2).Hyperlinks.Add(Anchor:=Cells(sheetObj.Index + 1, 2), Address:="", SubAddress:="'" & sheetObj.Name & "'!A1")
            End If
          
            '// �V�[�g�ݒ�l�̏o��
            If maxRow * maxCol > 0 Then
                idx = 3
                For idxRow = 1 To maxRow
                    For idxCol = 1 To maxCol
                        resultSheet.Cells(sheetObj.Index + 1, idx).NumberFormat = sheetObj.Cells(idxRow, idxCol).NumberFormat
                        resultSheet.Cells(sheetObj.Index + 1, idx).Value = sheetObj.Cells(idxRow, idxCol).Value
                        idx = idx + 1
                    Next
                    Call statGauge.addValue(1)
                Next
            End If
        End If
    Next
  
    Call gsPageSetup_Lines(resultSheet, 1)
  
    Set statGauge = Nothing
    '// �ʃu�b�N�ɏo�͂����ۂɂ́A����Ƃ��ɕۑ������߂Ȃ�
    ActiveWorkbook.Saved = (cmbOutput.Value = "0")
    Application.ScreenUpdating = True
End Sub

'// ////////////////////////////////////////////////////////////////////////////
'// END
'// ////////////////////////////////////////////////////////////////////////////
