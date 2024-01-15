VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSheetManage 
   Caption         =   "�V�[�g����"
   ClientHeight    =   5055
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6360
   OleObjectBlob   =   "frmSheetManage.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "frmSheetManage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'// ////////////////////////////////////////////////////////////////////////////
'// �v���W�F�N�g   : �g���c�[��
'// �^�C�g��       : �V�[�g�ݒ�t�H�[��
'// ���W���[��     : frmSheetManage
'// ����           : �V�[�g�̏��������s��
'// ////////////////////////////////////////////////////////////////////////////
'// Copyright (c) by Koichiro.
'// ////////////////////////////////////////////////////////////////////////////
Option Explicit
Option Base 0


'// //////////////////////////////////////////////////////////////////
'// �C�x���g�F �t�H�[�� �A�N�e�B�u��
Private Sub UserForm_Activate()
    '// ���O�`�F�b�N�i�V�[�g�L���j
    If Not gfPreCheck() Then
        Call Me.Hide
        Exit Sub
    End If
End Sub


'// //////////////////////////////////////////////////////////////////
'// �C�x���g�F �t�H�[����������
Private Sub UserForm_Initialize()
    '// �R���{�{�b�N�X�ݒ�
    Call gsSetCombo(cmbTarget, CMB_SMG_TARGET, 0)
    Call gsSetCombo(cmbView, CMB_SMG_VIEW, 0)
    Call gsSetCombo(cmbZoom, CMB_SMG_ZOOM, 0)
    Call gsSetCombo(cmbFilter, CMB_SMG_FILTER, 0)
    
    '// �L���v�V�����ݒ�
    frmSheetManage.Caption = LBL_SMG_FORM
    ckbScroll.Caption = LBL_SMG_SCROLL
    ckbFontColor.Caption = LBL_SMG_FONT_COLOR
    ckbLink.Caption = LBL_SMG_HYPERLINK
    ckbComment.Caption = LBL_SMG_COMMENT
    ckbHeader.Caption = LBL_SMG_HEAD_FOOT
    ckbMargin.Caption = LBL_SMG_MARGIN
    ckdbPageBreak.Caption = LBL_SMG_PAGEBREAK
    fraPrintOpt.Caption = LBL_SMG_PRINT_OPT
    optPrintNone.Caption = LBL_SMG_PRINT_NONE
    optPrintNoZoom.Caption = LBL_SMG_PRINT_100
    optPrintVert.Caption = LBL_SMG_PRINT_HRZ
    optPrint1Page.Caption = LBL_SMG_PRINT_1_PAGE
    lblTarget.Caption = LBL_SMG_TARGET
    lblView.Caption = LBL_SMG_VIEW
    lblZoom.Caption = LBL_SMG_ZOOM
    lblFilter.Caption = LBL_SMG_AUTOFILTER
    cmdExecute.Caption = LBL_COM_EXEC
    cmdClose.Caption = LBL_COM_CLOSE
    cmdSelectAll.Caption = LBL_COM_CHECK_ALL
    cmdClear.Caption = LBL_COM_UNCHECK
End Sub


'// //////////////////////////////////////////////////////////////////
'// �C�x���g�F ���ׂđI���{�^�� �N���b�N��
Private Sub cmdSelectAll_Click()
    Call setCheckBoxes(True)
End Sub


'// //////////////////////////////////////////////////////////////////
'// �C�x���g�F �I�������{�^�� �N���b�N��
Private Sub cmdClear_Click()
    Call setCheckBoxes(False)
End Sub


'// //////////////////////////////////////////////////////////////////
'// �C�x���g�F ���s�{�^�� �N���b�N��
Private Sub cmdExecute_Click()
On Error GoTo ErrorHandler
    Dim wkBook      As Workbook
    Dim wkSheet     As Worksheet
    Dim FilePath    As String
    Dim FileName    As String
    Dim compFiles   As String   '// �����t�@�C���X�V���A���������t�@�C������ێ�
  
    '// Zoom�l�̃`�F�b�N
    If IsNull(cmbZoom.Value) Then
      If IsNumeric(cmbZoom.Text) Then
        If CInt(cmbZoom.Text) < 10 Or CInt(cmbZoom.Text) > 400 Then
          Call MsgBox(MSG_VAL_10_400, vbOKOnly, APP_TITLE)
          Exit Sub
        End If
      Else
          Call MsgBox(MSG_VAL_10_400, vbOKOnly, APP_TITLE)
        Exit Sub
      End If
    End If
  
    '// �u�b�N�̕ی�m�F
    Select Case cmbTarget.Value
        Case 0 '// ���݂̃V�[�g
            If ActiveSheet.ProtectContents Then
                Call MsgBox(MSG_SHEET_PROTECTED, vbOKOnly, APP_TITLE)
                Exit Sub
            End If
        Case 1 '// ���݂̃��[�N�u�b�N
            For Each wkSheet In Worksheets
                If wkSheet.ProtectContents Then
                    Call MsgBox(MSG_SHEETS_PROTECTED, vbOKOnly, APP_TITLE)
                    Exit Sub
                End If
            Next
    End Select
  
    Call gsSuppressAppEvents
    
    '// �����ΏۃR���{�̒l�ɂ����{�V�[�g | �u�b�N | �f�B���N�g���P��}�Ɏ��s
    Select Case cmbTarget.Value
        Case 0    '// ���݂̃V�[�g
            Call psSetUpSheetProperty(ActiveSheet)
        Case 1    '// ���݂̃u�b�N
            For Each wkSheet In Worksheets
                Call psSetUpSheetProperty(wkSheet)
            Next
            Call ActiveWorkbook.Sheets(1).Activate
        Case 2    '// �f�B���N�g���P��
            '// ���s�O�ݒ�
            If Not gfShowSelectFolder(0, FilePath) Then
                Exit Sub
            End If
            
            FileName = Dir(FilePath & "\*.xls")
            Do While FileName <> BLANK
                Call Workbooks.Open(FilePath & "\" & FileName, ReadOnly:=False)
                For Each wkSheet In ActiveWorkbook.Worksheets
                    Call psSetUpSheetProperty(wkSheet)
                Next
                Call ActiveWorkbook.Sheets(1).Activate
                compFiles = compFiles & ActiveWorkbook.Name & Chr(10)
                Call ActiveWorkbook.Close(SaveChanges:=True)
                FileName = Dir
            Loop
            
            '// xlsx�`���iExcel2007�ȏ�j�ւ̑Ή�
            FileName = Dir(FilePath & "\*.xlsx")
            Do While FileName <> BLANK
                Call Workbooks.Open(FilePath & "\" & FileName, ReadOnly:=False)
                For Each wkSheet In ActiveWorkbook.Worksheets
                    Call psSetUpSheetProperty(wkSheet)
                Next
                Call ActiveWorkbook.Sheets(1).Activate
                compFiles = compFiles & Chr(10) & ActiveWorkbook.Name
                Call ActiveWorkbook.Close(SaveChanges:=True)
                
                FileName = Dir
            Loop
    End Select
    
    Call gsResumeAppEvents
    Call MsgBox(MSG_FINISHED, vbOKOnly, APP_TITLE)
    
    Call Me.Hide
    Exit Sub
  
ErrorHandler:
    Call gsResumeAppEvents
    Call gsShowErrorMsgDlg("frmSheetManage.cmdExecute_Click [" & ActiveWorkbook.FullName & "!" & ActiveSheet.Name & "]", Err, Nothing)
    Call MsgBox(MSG_COMPLETED_FILES & Chr(10) & compFiles, vbOKOnly, APP_TITLE)
End Sub


'// //////////////////////////////////////////////////////////////////
'// �C�x���g�F ����{�^�� �N���b�N��
Private Sub cmdClose_Click()
    Call Me.Hide
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// ���\�b�h�F   �`�F�b�N�{�b�N�X�ݒ�
'// �����F       �`�F�b�N�{�b�N�X�̐ݒ�������̐^�U�l�ɐݒ肷��B
'// �����F       setVal: �ݒ�l
'// �߂�l�F     �Ȃ�
'// ////////////////////////////////////////////////////////////////////////////
Private Sub setCheckBoxes(setVal As Boolean)
    ckbScroll.Value = setVal
    ckbFontColor.Value = setVal
    ckbLink.Value = setVal
    ckbComment.Value = setVal
    ckbHeader.Value = setVal
    ckbMargin.Value = setVal
    ckdbPageBreak.Value = setVal
    cmbView.Value = IIf(setVal, cmbView.Value, 0)
  
    '// �ȉ��̍��ڂɂ��Ắu�I������(setVal=false)�v�̏ꍇ�̂ݕ␳
    If setVal = False Then
        cmbZoom.Value = 0
        optPrintNone.Value = True
        cmbFilter.Value = 0
    End If
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// ���\�b�h�F   �V�[�g�ݒ�
'// �����F       �����̃V�[�g�̐��`�������s��
'//              ���s�{�^���N���b�N�C�x���g����Ăяo�����
'// �����F       wkSheet: �Ώۃ��[�N�V�[�g
'// ////////////////////////////////////////////////////////////////////////////
Private Sub psSetUpSheetProperty(wkSheet As Worksheet)
    Call wkSheet.Activate
    Call wkSheet.Cells(1, 1).Select  '// �O���t�Ȃǂ��A�N�e�B�u�ȏꍇ�̃G���[��������邽�߁AA1��I����Ԃɂ���
    Application.StatusBar = "Setting up: [" & wkSheet.Parent.Name & "!" & wkSheet.Name & "]"
    
    '// �r���[��ݒ� ���X�N���[���ݒ������Ƀr���[��ύX����K�v����iFreezePanes�ݒ�ŃG���[�ƂȂ邽�߁j
    Select Case cmbView.Value
        Case 1
            ActiveWindow.View = xlNormalView
        Case 2
            ActiveWindow.View = xlPageBreakPreview
    End Select
    
    '// �X�N���[���̏�����
    If ckbScroll.Value Then
        ActiveWindow.FreezePanes = False
        ActiveWindow.Split = False
        ActiveWindow.ScrollColumn = 1
        ActiveWindow.ScrollRow = 1
    End If
    
    '// �Y�[����������
    If IsNull(cmbZoom.Value) Or cmbZoom.Value <> 0 Then
        ActiveWindow.Zoom = cmbZoom.Text
    End If
    
    '// �I�[�g�t�B���^�̐ݒ� �u0:�w�薳���v�͖���
    Select Case cmbFilter.Value
        Case 1 '// �t�B���^����
            If wkSheet.AutoFilterMode Then
                Call wkSheet.Cells.AutoFilter
            End If
        Case 2 '// �S�ĕ\��
            If WorksheetFunction.CountA(ActiveSheet.UsedRange) > 1 Then
                Call wkSheet.ShowAllData
            End If
        Case 3 '// �P�s�ڂŃt�B���^
            If Not wkSheet.AutoFilterMode And WorksheetFunction.CountA(ActiveSheet.UsedRange) > 1 Then
                Call wkSheet.Cells.AutoFilter
            End If
    End Select
    
    '// �n�C�p�[�����N���폜
    If ckbLink.Value Then
        Call wkSheet.Hyperlinks.Delete
    End If
    
    '// ����̊g��/�k��
    With wkSheet.PageSetup
        If optPrintNoZoom.Value Then
            .Zoom = 100
        ElseIf optPrintVert.Value Then
            .Zoom = False
            .FitToPagesWide = 1
            .FitToPagesTall = False
        ElseIf optPrint1Page.Value Then
            .Zoom = False
            .FitToPagesWide = 1
            .FitToPagesTall = 1
        End If
    End With
    
    '// �w�b�_�ƃt�b�^�̐ݒ�
    If ckbHeader.Value Then
        Call mdlCommon.gsPageSetup_Header(wkSheet)
    End If
    
    '// �}�[�W���̐ݒ�
    If ckbMargin.Value Then
        Call mdlCommon.gsPageSetup_Margin(wkSheet)
    End If
    
    '// ���y�[�W�ƈ���͈͂�����
    If ckdbPageBreak.Value Then
        Call ActiveSheet.ResetAllPageBreaks
        wkSheet.PageSetup.PrintArea = ""
    End If
    
    '// �t�H���g�F�̏�����
    If ckbFontColor.Value Then
        wkSheet.Cells.Font.ColorIndex = xlAutomatic
    End If
    
    '// �R�����g���폜
    If ckbComment.Value Then
        Call wkSheet.Cells.ClearComments
    End If
    
    Call wkSheet.Cells(1, 1).Select
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// END
'// ////////////////////////////////////////////////////////////////////////////
