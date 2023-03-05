VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCompSheet 
   Caption         =   "�V�[�g��r"
   ClientHeight    =   4905
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6600
   OleObjectBlob   =   "frmCompSheet.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "frmCompSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'// ////////////////////////////////////////////////////////////////////////////
'// �v���W�F�N�g   : �g���c�[��
'// �^�C�g��       : �V�[�g��r�t�H�[��
'// ���W���[��     : frmCompSheet
'// ����           : �V�[�g�̔�r���s��
'//                  Excel2013����W���@�\�Ŏ������ꂽ���߁A���̋@�\�̃����e�i���X�͍���s��Ȃ��B
'// ////////////////////////////////////////////////////////////////////////////
'// Copyright (c) by Koichiro.
'// ////////////////////////////////////////////////////////////////////////////
Option Explicit
Option Base 0

'// ////////////////////////////////////////////////////////////////////////////
'// �v���C�x�[�g�萔
Private Const FLG_INS_ROW               As String = "$ins_extools"
Private Const FLG_DEL_ROW               As String = "$del_extools"
Private Const CLR_DIFF_INS_ROW          As Integer = 34  '// 42
Private Const CLR_DIFF_DEL_ROW          As Integer = 15  '// 48
Private Const STATINTERVAL              As Long = 100    '// �X�e�[�^�X�o�[�̍X�V�C���^�[�o��(�P��:�s)

'// ////////////////////////////////////////////////////////////////////////////
'// �v���C�x�[�g�ϐ�
'// �����^�C�v
Private Type udDiff
  sheet As String
  Row   As Integer
  Col   As Integer
  val_1 As String
  val_2 As String
  note  As String
End Type

'// �s�K���^�C�v
Private Type udRowPair
  row1  As Long
  row2  As Long
End Type

Private pDiff()                 As udDiff       '// �����̌���
Private pMatched()              As udRowPair    '// �K���s�̈ꗗ


'// //////////////////////////////////////////////////////////////////
'// �C�x���g�F �t�H�[�� �A�N�e�B�u��
Private Sub UserForm_Activate()
    '// ���O�`�F�b�N�i�V�[�g�L���j
    If Not gfPreCheck() Then
        Call Me.Hide
        Exit Sub
    End If
    
  Call psSetSheetCombo(cmbSheet_1.Name)
  Call psSetSheetCombo(cmbSheet_2.Name)
End Sub


'// //////////////////////////////////////////////////////////////////
'// �C�x���g�F �t�H�[�� ��������
Private Sub UserForm_Initialize()
  '// �R���{�{�b�N�X�ݒ�
  Call gsSetCombo(cmbResultType, CMB_CMP_MARKER, 0)
  Call gsSetCombo(cmbCompareMode, CMB_CMP_METHOD, 0)
  Call gsSetCombo(cmbOutput, CMB_CMP_OUTPUT, 0)
  
  '// �L���v�V�����ݒ�
  frmCompSheet.Caption = LBL_CMP_FORM
  cmdExecute.Caption = LBL_COM_EXEC
  cmdClose.Caption = LBL_COM_CLOSE
  cmdFile_1.Caption = LBL_COM_BROWSE
  cmdFile_2.Caption = LBL_COM_BROWSE
  mpgTarget.Pages(0).Caption = LBL_CMP_MODE_SHEET
  mpgTarget.Pages(1).Caption = LBL_CMP_MODE_BOOK
  ckbShowComments.Caption = LBL_CMP_SHOW_COMMENT
  fraOption.Caption = LBL_CMP_OPTIONS
  lblOriginalSheet.Caption = LBL_CMP_SHEET1
  lblTargetSheet.Caption = LBL_CMP_SHEET2
  lblOriginalBook.Caption = LBL_CMP_BOOK1
  lblTargetBook.Caption = LBL_CMP_BOOK2
  lblOutput.Caption = LBL_CMP_RESULT
  lblMarker.Caption = LBL_CMP_MARKER
  lblMethod.Caption = LBL_CMP_METHOD
End Sub


'// //////////////////////////////////////////////////////////////////
'// �C�x���g�F ����{�^�� �N���b�N��
Private Sub cmdClose_Click()
  Call Me.Hide
End Sub


'// //////////////////////////////////////////////////////////////////
'// �C�x���g�F �Q�ƃ{�^�� �N���b�N��
Private Sub cmdFile_1_Click()
  txtFileName_1.Text = pfGetFileName(txtFileName_1.Text)
End Sub


'// //////////////////////////////////////////////////////////////////
'// �C�x���g�F �Q�ƃ{�^�� �N���b�N��
Private Sub cmdFile_2_Click()
  txtFileName_2.Text = pfGetFileName(txtFileName_2.Text)
End Sub


'// //////////////////////////////////////////////////////////////////
'// �C�x���g�F ���s�{�^�� �N���b�N��
Private Sub cmdExecute_Click()
  Select Case mpgTarget.Value
    Case 0
      Call psCompSheet
    Case 1
      Call psCompBook
  End Select
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// ���\�b�h�F   �V�[�g��r
'// �����F       �V�[�g�̔�r���s��
'// �����F       �Ȃ�
'// ////////////////////////////////////////////////////////////////////////////
Private Sub psCompSheet()
On Error GoTo ErrorHandler
  Dim errCnt    As Long   '// �G���[��
  
  Application.ScreenUpdating = False
  '// ��r���s
  Erase pDiff
  Call psExecComp(Worksheets(cmbSheet_1.Text), Worksheets(cmbSheet_2.Text), cmbResultType.Value, cmbCompareMode.Value, errCnt)
    
  '// ���ʂ̏o��
  If errCnt > 0 Then
    Call psShowResult(ActiveWorkbook)
  Else
    '// �[�����̏ꍇ�̓��b�Z�[�W��\��
    Call MsgBox(MSG_NO_DIFF, vbOKOnly, APP_TITLE)
  End If
  
  Call Me.Hide
  Application.ScreenUpdating = True
  Exit Sub
  
ErrorHandler:
  Call gsShowErrorMsgDlg("frmCompSheet.psCompSheet", Err)
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// ���\�b�h�F   �u�b�N��r
'// �����F       �u�b�N�̔�r���s��
'// �����F       �Ȃ�
'// ////////////////////////////////////////////////////////////////////////////
Private Sub psCompBook()
On Error GoTo ErrorHandler
  Dim errCnt    As Long       '// �G���[��
  Dim idx       As Integer    '// �V�[�g�̃C���f�N�X
  Dim book1     As Workbook   '// �Ώۃu�b�N�i�P�j
  Dim book2     As Workbook   '// �Ώۃu�b�N�i�Q�j
  Dim cntRows   As Long
  Dim cntCols   As Integer
  
  '// ���̓`�F�b�N
  If (Trim(txtFileName_1.Text) = BLANK) Or (Trim(txtFileName_2.Text) = BLANK) Then
    Call MsgBox(MSG_ERROR_NEED_BOOKNAME, vbOKOnly, APP_TITLE)
    If Trim(txtFileName_1.Text) = BLANK Then
      Call txtFileName_1.SetFocus
    Else
      Call txtFileName_2.SetFocus
    End If
    Exit Sub
  End If

  Application.ScreenUpdating = False
  '// �V�[�g���J��
  Set book1 = pfGetBook(txtFileName_1.Text)
  Set book2 = pfGetBook(txtFileName_2.Text)
  
  If (book1 Is Nothing) Or (book2 Is Nothing) Then
    Application.ScreenUpdating = True
    Exit Sub
  End If
  
  '// ��r���s�i�V�[�g�\���j
  Erase pDiff
  errCnt = 0
  For idx = 1 To book1.Worksheets.Count
    '// �z��ւ̊i�[
    ReDim Preserve pDiff(errCnt + 1)
    pDiff(errCnt).val_1 = "�V�[�g�F " & book1.Worksheets(idx).Name
    If idx <= book2.Worksheets.Count Then
      pDiff(errCnt).val_2 = "�V�[�g�F " & book2.Worksheets(idx).Name
    End If
    '// �V�[�g�����Ⴄ�ꍇ
    If pDiff(errCnt).val_1 <> pDiff(errCnt).val_2 Then
      pDiff(errCnt).sheet = book2.Worksheets(errCnt + 1).Name
      pDiff(errCnt).Row = 1
      pDiff(errCnt).Col = 1
      pDiff(errCnt).note = MSG_SHEET_NAME
      errCnt = errCnt + 1
    End If
  Next
  
  '// �V�[�g�\�����قȂ�ꍇ�ɂ͏I��
  If errCnt > 0 Then
    Call MsgBox(MSG_UNMATCH_SHEET, vbOKOnly, APP_TITLE)
    Call psShowResult(book2)
    Call Me.Hide
    Set book1 = Nothing
    Set book2 = Nothing
    Application.ScreenUpdating = True
    Exit Sub
  End If
  
  '// ��r���s�i�V�[�g���j
  Erase pDiff
  errCnt = 0
  For idx = 1 To book1.Worksheets.Count
    Call psExecComp(book1.Worksheets(idx), book2.Worksheets(idx), cmbResultType.Value, cmbCompareMode.Value, errCnt)
  Next
  
  If errCnt > 0 Then
    '// ���ʂ̏o��
    Call psShowResult(book2)
    Call Me.Hide
  Else
    '// �[�����̏ꍇ�̓��b�Z�[�W��\��
    Call MsgBox(MSG_NO_DIFF, vbOKOnly, APP_TITLE)
  End If
  
  Set book1 = Nothing
  Set book2 = Nothing
  Application.ScreenUpdating = True
  Exit Sub
  
ErrorHandler:
  Call gsShowErrorMsgDlg("frmCompBook.psCompBook", Err)
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// ���\�b�h�F   �V�[�g�̃R���{�ݒ�
'// �����F       �w�肳�ꂽ�u�b�N�Ɋ܂܂��V�[�g���������A�R���{�{�b�N�X�ɐݒ肷��
'// �����F       comboName: �R���{�{�b�N�X����
'// ////////////////////////////////////////////////////////////////////////////
Private Sub psSetSheetCombo(comboName As String)
  Dim combo         As ComboBox
  Dim wkSheet       As Worksheet
  Dim currentSheet  As String  '// �R���{�{�b�N�X�������O�̃V�[�g��
  Dim defaultIdx    As Integer
  
  '// ������
  defaultIdx = 0
  Set combo = Me.Controls(comboName)
  currentSheet = combo.Text
  Call combo.Clear
  
  '// �V�[�g���擾
  For Each wkSheet In ActiveWorkbook.Worksheets
    Call combo.AddItem(wkSheet.Name)
    If wkSheet.Name = currentSheet Then
      defaultIdx = combo.ListCount - 1
    End If
  Next
  
  combo.ListIndex = defaultIdx
  Set combo = Nothing
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// ���\�b�h�F   �t�@�C�����擾
'// �����F       �_�C�A���O��\�����A�t�@�C������Ԃ��B
'// �����F       defaultVal: �t�@�C����
'// �߂�l�F     �t���p�X�̃t�@�C����
'// ////////////////////////////////////////////////////////////////////////////
Private Function pfGetFileName(defaultVal As String)
'  Call gsCheckVersion
  
  pfGetFileName = Application.GetOpenFilename(FileFilter:=Replace(APP_EXL_FILE, "#", EXCEL_FILE_EXT))
  If pfGetFileName = False Then
    pfGetFileName = defaultVal
  End If
End Function


'// ////////////////////////////////////////////////////////////////////////////
'// ���\�b�h�F   �u�b�N�̃I�[�v��
'// �����F       �u�b�N�I�u�W�F�N�g��Ԃ�
'// �����F       fileName: �t�@�C����
'// �߂�l�F     �u�b�N�I�u�W�F�N�g
'// ////////////////////////////////////////////////////////////////////////////
Private Function pfGetBook(FileName As String) As Workbook
On Error GoTo ErrorHandler
  Set pfGetBook = Workbooks.Open(FileName:=FileName, ReadOnly:=False)
  Exit Function

ErrorHandler:
  Call MsgBox(MSG_NO_FILE & " [" & FileName & " ]", vbOKOnly, APP_TITLE)
  Set pfGetBook = Nothing
End Function


'// ////////////////////////////////////////////////////////////////////////////
'// ���\�b�h�F   ��r
'// �����F       ��r���s��
'// �����F       wkSheet1, wkSheet2: ��r�ΏۃV�[�g
'//              compResult �o�͌`�� 0:�������Ȃ�  1:�����𒅐F  2:�Z���𒅐F
'//              compMode ��r���[�h 0:�e�L�X�g  1:�l  2:�l�܂��̓e�L�X�g
'//              cnt: �������i�Ăяo�������p���j
'// �߂�l�F     �Ȃ�
'// ////////////////////////////////////////////////////////////////////////////
Private Sub psExecComp(wkSheet1 As Worksheet, wkSheet2 As Worksheet, _
                       compResult As Integer, compMode As Integer, ByRef cnt As Long)
On Error GoTo ErrorHandler
  Dim idxRow          As Long           '// ���o���̍s�C���f�N�X
  Dim idxCol          As Integer        '// ���o���̗�C���f�N�X
  Dim tRange          As udTargetRange  '// ���o�͈�
  Dim isDiff_v        As Boolean        '// �l�̈Ⴂ�L������
  Dim isDiff_t        As Boolean        '// �e�L�X�g�̈Ⴂ�L������
  Dim isDiff_f        As Boolean        '// �����̈Ⴂ�L������
  Dim isDiff          As Boolean        '// �����L���̑�������
  
  '// �����͈͂̐ݒ�
  tRange.minRow = IIf(wkSheet1.UsedRange.Row < wkSheet2.UsedRange.Row, wkSheet1.UsedRange.Row, wkSheet2.UsedRange.Row)
  tRange.minCol = IIf(wkSheet1.UsedRange.Column < wkSheet2.UsedRange.Column, wkSheet1.UsedRange.Column, wkSheet2.UsedRange.Column)
  tRange.maxRow = IIf((wkSheet1.UsedRange.Row + wkSheet1.UsedRange.Rows.Count) > (wkSheet2.UsedRange.Row + wkSheet2.UsedRange.Rows.Count), (wkSheet1.UsedRange.Row + wkSheet1.UsedRange.Rows.Count - 1), (wkSheet2.UsedRange.Row + wkSheet2.UsedRange.Rows.Count - 1))
  tRange.maxCol = IIf((wkSheet1.UsedRange.Column + wkSheet1.UsedRange.Columns.Count) > (wkSheet2.UsedRange.Column + wkSheet2.UsedRange.Columns.Count), (wkSheet1.UsedRange.Column + wkSheet1.UsedRange.Columns.Count - 1), (wkSheet2.UsedRange.Column + wkSheet2.UsedRange.Columns.Count - 1))
  
  '// �e�V�[�g���Q�s�ȏ�A�����v�T�s�ȏ�̏ꍇ�̂݁A�s�����̐��������{�i�s�����Ȃ��ꍇ�͐����ł̗�O�������ʓ|�Ȃ��߁j
  If wkSheet1.UsedRange.Rows.Count > 1 And wkSheet2.UsedRange.Rows.Count > 1 And wkSheet1.UsedRange.Rows.Count + wkSheet2.UsedRange.Rows.Count > 4 Then
    Call psPadRowDiff(compMode, wkSheet1, wkSheet2, tRange.maxRow, tRange.maxCol)
  End If
  
  '// �����͈͂̐ݒ�
  tRange.maxRow = IIf((wkSheet1.UsedRange.Row + wkSheet1.UsedRange.Rows.Count) > (wkSheet2.UsedRange.Row + wkSheet2.UsedRange.Rows.Count), (wkSheet1.UsedRange.Row + wkSheet1.UsedRange.Rows.Count - 1), (wkSheet2.UsedRange.Row + wkSheet2.UsedRange.Rows.Count - 1))
  
  
  '// ��r
  Call wkSheet2.Activate
  For idxRow = tRange.minRow To tRange.maxRow
    If wkSheet2.Cells(idxRow, 1).NoteText = FLG_INS_ROW Then
      wkSheet2.Cells(idxRow, 1).NoteText (MSG_INS_ROW)
      '// �z��ւ̊i�[
      ReDim Preserve pDiff(cnt + 1)
      pDiff(cnt).sheet = wkSheet2.Name
      pDiff(cnt).Row = idxRow
      pDiff(cnt).Col = 1
      pDiff(cnt).note = MSG_INS_ROW
      cnt = cnt + 1
    ElseIf wkSheet2.Cells(idxRow, 1).NoteText = FLG_DEL_ROW Then
      wkSheet2.Cells(idxRow, 1).NoteText (MSG_DEL_ROW)
      '// �z��ւ̊i�[
      ReDim Preserve pDiff(cnt + 1)
      pDiff(cnt).sheet = wkSheet2.Name
      pDiff(cnt).Row = idxRow
      pDiff(cnt).Col = 1
      pDiff(cnt).note = MSG_DEL_ROW
      cnt = cnt + 1
    Else
      For idxCol = tRange.minCol To tRange.maxCol
        isDiff_v = False
        isDiff_t = False
        isDiff_f = False
        '// �e�L�X�g�̈Ⴂ���m�F
        isDiff_t = (wkSheet1.Cells(idxRow, idxCol).Text <> wkSheet2.Cells(idxRow, idxCol).Text)
        '// �l�̈Ⴂ���m�F
        If IsError(wkSheet1.Cells(idxRow, idxCol)) And IsError(wkSheet2.Cells(idxRow, idxCol)) Then
          isDiff_v = False
        ElseIf IsError(wkSheet1.Cells(idxRow, idxCol)) Xor IsError(wkSheet2.Cells(idxRow, idxCol)) Then
          isDiff_v = True
        Else
          isDiff_v = (wkSheet1.Cells(idxRow, idxCol).Value <> wkSheet2.Cells(idxRow, idxCol).Value)
        End If
        '// �����̈Ⴂ���m�F
        isDiff_f = (CStr(wkSheet1.Cells(idxRow, idxCol).NumberFormat) <> CStr(wkSheet2.Cells(idxRow, idxCol).NumberFormat))
        
        isDiff = (isDiff_t And compMode = 0) Or (isDiff_v And compMode = 1) Or ((compMode = 2) And (isDiff_v And isDiff_t))
        If isDiff Then
          '// ���F�i�w�莞�j
          If Not wkSheet2.ProtectContents Then
            Select Case compResult
              Case 1   '// �����𒅐F
                wkSheet2.Cells(idxRow, idxCol).Font.ColorIndex = COLOR_DIFF_CELL
              Case 2   '// �Z���𒅐F
                wkSheet2.Cells(idxRow, idxCol).Interior.ColorIndex = COLOR_DIFF_CELL
              Case 3   '// �g�𒅐F
                wkSheet2.Cells(idxRow, idxCol).Borders(xlEdgeLeft).LineStyle = xlContinuous
                wkSheet2.Cells(idxRow, idxCol).Borders(xlEdgeLeft).ColorIndex = COLOR_DIFF_CELL
                wkSheet2.Cells(idxRow, idxCol).Borders(xlEdgeTop).LineStyle = xlContinuous
                wkSheet2.Cells(idxRow, idxCol).Borders(xlEdgeTop).ColorIndex = COLOR_DIFF_CELL
                wkSheet2.Cells(idxRow, idxCol).Borders(xlEdgeBottom).LineStyle = xlContinuous
                wkSheet2.Cells(idxRow, idxCol).Borders(xlEdgeBottom).ColorIndex = COLOR_DIFF_CELL
                wkSheet2.Cells(idxRow, idxCol).Borders(xlEdgeRight).LineStyle = xlContinuous
                wkSheet2.Cells(idxRow, idxCol).Borders(xlEdgeRight).ColorIndex = COLOR_DIFF_CELL
            End Select
          End If
          '// �R�����g
          If Not wkSheet2.ProtectContents Then
            Select Case compMode
              Case 0   '// �e�L�X�g
                Call wkSheet2.Cells(idxRow, idxCol).NoteText(IIf(wkSheet1.Cells(idxRow, idxCol).Text = BLANK, "<Blank>", wkSheet1.Cells(idxRow, idxCol).Text))
              Case 1   '// �l
                If IsError(wkSheet1.Cells(idxRow, idxCol)) Then
                  Call wkSheet2.Cells(idxRow, idxCol).NoteText(wkSheet1.Cells(idxRow, idxCol).Text)
                Else
                  Call wkSheet2.Cells(idxRow, idxCol).NoteText(IIf(wkSheet1.Cells(idxRow, idxCol).Value = BLANK, "<Blank>", wkSheet1.Cells(idxRow, idxCol).Value))
                End If
              Case 2   '// �e�L�X�g�܂��͒l
                If IsError(wkSheet1.Cells(idxRow, idxCol)) Then
                  Call wkSheet2.Cells(idxRow, idxCol).NoteText(wkSheet1.Cells(idxRow, idxCol).Text)
                Else
                  Call wkSheet2.Cells(idxRow, idxCol).NoteText("�y÷�āz" & wkSheet1.Cells(idxRow, idxCol).Text & vbLf & "�y�l�z" & wkSheet1.Cells(idxRow, idxCol).Value)
                End If
            End Select
          End If
          
          '// �z��ւ̊i�[
          ReDim Preserve pDiff(cnt + 1)
          pDiff(cnt).sheet = wkSheet2.Name
          pDiff(cnt).Row = idxRow
          pDiff(cnt).Col = idxCol
          '// �l���Ⴄ�ꍇ
          If isDiff_v Or isDiff_t Then
            pDiff(cnt).val_1 = IIf(IsError(wkSheet1.Cells(idxRow, idxCol)), wkSheet1.Cells(idxRow, idxCol).Text, wkSheet1.Cells(idxRow, idxCol).Value)
            pDiff(cnt).val_2 = IIf(IsError(wkSheet2.Cells(idxRow, idxCol)), wkSheet2.Cells(idxRow, idxCol).Text, wkSheet2.Cells(idxRow, idxCol).Value)
          End If
          
          '// �������Ⴄ�ꍇ
          If isDiff_f Then
            pDiff(cnt).note = "����"
          End If
          
          cnt = cnt + 1
          '// �L�[����
          If GetAsyncKeyState(27) <> 0 Then
            Application.StatusBar = False
            Exit Sub
          End If
        End If
      Next
    End If
    
    '// �X�e�[�^�X�o�[�X�V
    If idxRow Mod STATINTERVAL = 0 Then
      Application.StatusBar = "�Z�����e��r��... [ �s: " & CStr(idxRow) & " / ����: " & CStr(cnt) & " ]" & IIf(wkSheet1.Name = wkSheet2.Name, wkSheet1.Name, BLANK)
    End If
  Next
  
  Application.StatusBar = False
  Exit Sub
ErrorHandler:
  Call gsShowErrorMsgDlg("frmCompSheet.psExecComp", Err)
  Application.StatusBar = False
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// ���\�b�h�F   ��r���ʏo��
'// �����F       ��r���ʂ�ʃu�b�N�ŏo�͂���
'// �����F       �Ȃ�
'// �߂�l�F     �Ȃ�
'// �C�������F   �Ȃ�
'// ////////////////////////////////////////////////////////////////////////////
Private Sub psShowResult(wkBook As Workbook)
  Dim wkSheet   As Worksheet
  Dim idxRow    As Integer
  Dim needLink  As Boolean
  
  needLink = Not wkBook.MultiUserEditing And (cmbOutput.Value <> "0")
  
  '// �o�͐�̐ݒ�
  Select Case cmbOutput.Value
    Case "0"
      Call Workbooks.Add
      Set wkSheet = ActiveWorkbook.ActiveSheet
    Case "1"
      Set wkSheet = wkBook.Sheets.Add(After:=wkBook.Sheets(wkBook.Sheets.Count))
  End Select
  
  '// �w�b�_�̐ݒ�
  wkSheet.Cells(1, 1).Value = "�V�[�g"
  wkSheet.Cells(1, 2).Value = "�Z��"
  wkSheet.Cells(1, 3).Value = "��r���Ƃ̒l (" & IIf(mpgTarget.Value = 0, cmbSheet_1.Text, txtFileName_1.Text) & ")"
  wkSheet.Cells(1, 4).Value = "��r��̒l (" & IIf(mpgTarget.Value = 0, cmbSheet_2.Text, txtFileName_2.Text) & ")"
  wkSheet.Cells(1, 5).Value = "���l"
  
  wkSheet.Columns("C:D").NumberFormat = "@"
  
  '// �����̐ݒ�
  For idxRow = 0 To UBound(pDiff) - 1
    wkSheet.Cells(idxRow + 2, 1).Value = pDiff(idxRow).sheet
    wkSheet.Cells(idxRow + 2, 2).Value = mdlCommon.gfGetColIndexString(pDiff(idxRow).Col) & CStr(pDiff(idxRow).Row)
    If needLink Then
      Call wkSheet.Cells(idxRow + 2, 2).Hyperlinks.Add(Anchor:=Cells(idxRow + 2, 2), Address:=BLANK, SubAddress:="'" & pDiff(idxRow).sheet & "'!" & wkSheet.Cells(idxRow + 2, 2).Value)
    End If
    wkSheet.Cells(idxRow + 2, 3).Value = pDiff(idxRow).val_1
    wkSheet.Cells(idxRow + 2, 4).Value = pDiff(idxRow).val_2
    wkSheet.Cells(idxRow + 2, 5).Value = pDiff(idxRow).note
  Next
  
  '// //////////////////////////////////////////////////////
  '// �����̐ݒ�
  '// ���̐ݒ�
  wkSheet.Columns("A").ColumnWidth = 10
  wkSheet.Columns("B").ColumnWidth = 8
  wkSheet.Columns("C:E").ColumnWidth = 30
  
  '// �g���̐ݒ�
  Call wkSheet.Range(wkSheet.Cells(1, 1), wkSheet.Cells(UBound(pDiff) + 1, 5)).Select
  Call gsDrawLine_Data
  
  '// �w�b�_�̏C��
  Call wkSheet.Range("A1:E1").Select
  Call gsDrawLine_Header
  
  '//�t�H���g
  wkSheet.Cells.Select
  Selection.Font.Name = APP_FONT
  Selection.Font.Size = APP_FONT_SIZE
  Call wkSheet.Cells(1, 1).Select
  
  '// ����Ƃ��ɕۑ������߂Ȃ�
  If cmbOutput.Value = "0" Then
    ActiveWorkbook.Saved = True
  End If
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// ���\�b�h�F   �s�ړ��i����s�ł���ꍇ�ɉ��ʍs�ֈړ��j
'// �����F       �n�m�c�̃T�u���\�b�h
'// �����F       compMode ��r���[�h 0:�e�L�X�g  1:�l  2:�l�܂��̓e�L�X�g
'//              wkSheet1, wkSheet2: ��r�ΏۃV�[�g
'//              idxRow1, idxRow2 �ΏۃV�[�g���̔�r�Ώۍs
'//              maxCol: ��r��
'// �߂�l�F     �V�[�g�Q�̈�v�s�ԍ�
'// ////////////////////////////////////////////////////////////////////////////
Private Function pfSnake(compMode As Integer, wkSheet1 As Worksheet, wkSheet2 As Worksheet, _
                         idxRow1 As Long, idxRow2 As Long, _
                         maxRow As Long, maxCol As Integer) As Long
  If idxRow1 < 1 Or idxRow2 < 1 Then
    pfSnake = idxRow2
    Exit Function
  End If
  
  Do While (idxRow1 < maxRow) And (idxRow2 < maxRow) And (pfGetRowScore(compMode, wkSheet1, wkSheet2, idxRow1, idxRow2, maxCol, False) > 0)

    
    ReDim Preserve pMatched(UBound(pMatched) + 1)
    pMatched(UBound(pMatched)).row1 = idxRow1
    pMatched(UBound(pMatched)).row2 = idxRow2
    
    idxRow1 = idxRow1 + 1
    idxRow2 = idxRow2 + 1
    
'    If (idxRow1 Mod STATINTERVAL = 0) Or (idxRow2 Mod STATINTERVAL = 0) Then
'      Application.StatusBar = "�s�������͒�... [ ��r��: " & CStr(idxRow1) & " / ��r��: " & CStr(idxRow2) & " ]" & IIf(wkSheet1.Name = wkSheet2.Name, wkSheet1.Name, BLANK)
'    End If
  Loop
  
  Application.StatusBar = False
  pfSnake = idxRow2
End Function


'// ////////////////////////////////////////////////////////////////////////////
'// ���\�b�h�F   �s��r
'// �����F       �w��s������ł��邩��^�U�l�ŕԂ�
'// �����F       compMode ��r���[�h 0:�e�L�X�g  1:�l  2:�l�܂��̓e�L�X�g
'//              wkSheet1, wkSheet2: ��r�ΏۃV�[�g
'//              idxRow1, idxRow2 �ΏۃV�[�g���̔�r�Ώۍs
'//              maxCol: ��r��
'//              getScore: true:�X�R�A���擾����, false:�^�U�l�Ƃ��� 0 �܂���1��Ԃ��B
'// �߂�l�F     0�`1�̎����B 0�͍s�����S�Ɉ�v���Ȃ��B1�͍s�����S�Ɉ�v����
'// ////////////////////////////////////////////////////////////////////////////
Private Function pfGetRowScore(compMode As Integer, wkSheet1 As Worksheet, wkSheet2 As Worksheet, idxRow1 As Long, idxRow2 As Long, maxCol As Integer, getScore As Boolean) As Double
On Error GoTo ErrorHandler
  Dim idxCol    As Integer
  Dim isDiff_v  As Boolean        '// �l�̈Ⴂ�L������
  Dim isDiff_t  As Boolean        '// �e�L�X�g�̈Ⴂ�L������
  Dim isDiff_f  As Boolean        '// �����̈Ⴂ�L������
  Dim score     As Long           '// �����Ή��F�X�R�A�ɂ��s�ގ��x�̗D�򔻒�
  
  For idxCol = 1 To maxCol
    isDiff_v = False
    isDiff_t = False
    isDiff_f = False
    '// �e�L�X�g�̈Ⴂ���m�F
    isDiff_t = (wkSheet1.Cells(idxRow1, idxCol).Text <> wkSheet2.Cells(idxRow2, idxCol).Text)
    '// �l�̈Ⴂ���m�F
    isDiff_v = (wkSheet1.Cells(idxRow1, idxCol).Value <> wkSheet2.Cells(idxRow2, idxCol).Value)
    '// �����̈Ⴂ���m�F
    isDiff_f = (CStr(wkSheet1.Cells(idxRow1, idxCol).NumberFormat) <> CStr(wkSheet2.Cells(idxRow2, idxCol).NumberFormat))
    
    If (isDiff_t And compMode = 0) Or (isDiff_v And compMode = 1) Or ((compMode = 2) And (isDiff_v And isDiff_t)) Then
      If Not getScore Then
        pfGetRowScore = 0
        Exit Function
      End If
    Else
      score = score + 1
    End If
  Next
  
  pfGetRowScore = CDbl(score / idxCol)
  Exit Function

ErrorHandler:
  pfGetRowScore = 0
End Function


'// ////////////////////////////////////////////////////////////////////////////
'// ���\�b�h�F   �G�f�B�b�g�O���t����
'// �����F       O(ND) �A���S���Y���ł̑������s��
'// �����F       compMode ��r���[�h 0:�e�L�X�g  1:�l  2:�l�܂��̓e�L�X�g
'//              wkSheet1, wkSheet2: ��r�ΏۃV�[�g
'//              maxRow, maxCol: ��r��
'// �߂�l�F     true:�s�͓���  false:�s�͕s��v
'// ////////////////////////////////////////////////////////////////////////////
Public Sub psOnd(compMode As Integer, wkSheet1 As Worksheet, wkSheet2 As Worksheet, _
                         maxRow As Long, maxCol As Integer)
  Dim currentIdx  As udRowPair
  Dim offset      As Long
  Dim idxSed      As Long
  Dim idxRoute    As Long
  Dim aryTemp()   As Long
  
  currentIdx.row1 = 0
  currentIdx.row2 = 0
  offset = maxRow
  
  ReDim pMatched(0)
  ReDim aryTemp(maxRow * 2)
  
  aryTemp(1 + offset) = 0
  
  For idxSed = 0 To maxRow * 2 'M + N
    For idxRoute = (-1 * idxSed) To idxSed Step 2
      If (idxRoute = -1 * idxSed) Then
        currentIdx.row2 = aryTemp(idxRoute + 1 + offset)
      ElseIf ((idxRoute <> idxSed) And aryTemp(idxRoute - 1 + offset) < aryTemp(idxRoute + 1 + offset)) Then
        currentIdx.row2 = aryTemp(idxRoute + 1 + offset)
      Else
        currentIdx.row2 = aryTemp(idxRoute - 1 + offset) + 1
      End If
      
      currentIdx.row1 = currentIdx.row2 - idxRoute
      
      aryTemp(idxRoute + offset) = pfSnake(compMode, wkSheet1, wkSheet2, currentIdx.row1, currentIdx.row2, maxRow, maxCol)
      If (currentIdx.row1 >= maxRow Or currentIdx.row2 >= maxRow) Then
        Exit Sub
      End If
    Next
  Next
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// ���\�b�h�F   �s�����␳
'// �����F       �s�̐������s������A�����s��}������
'// �����F       compMode ��r���[�h 0:�e�L�X�g  1:�l  2:�l�܂��̓e�L�X�g
'//              wkSheet1, wkSheet2: ��r�ΏۃV�[�g
'//              maxRow, maxCol: ��r�s���A��
'// ////////////////////////////////////////////////////////////////////////////
Public Sub psPadRowDiff(compMode As Integer, wkSheet1 As Worksheet, wkSheet2 As Worksheet, _
                        maxRow As Long, maxCol As Integer)
  Dim idxMatched  As Long
  Dim pad1        As Long
  Dim pad2        As Long
  Dim rowDiff1    As Long
  Dim rowDiff2    As Long
  Dim rowDiff     As Long  '// �Ō�̍s��␳����ۂ̍���
  Dim cnt         As Long
  Dim rowCntLimit As Long  '// �s�ǉ��̏���l�i������Ԃ̂Q�V�[�g�̍s���̍��v�j
  
  Call psOnd(compMode, wkSheet1, wkSheet2, maxRow, maxCol)
  
  rowCntLimit = wkSheet1.UsedRange.Rows.Count + wkSheet2.UsedRange.Rows.Count
  ReDim pDiffRows(0)
  idxMatched = 1
  Do
    If idxMatched > UBound(pMatched) Then
      Exit Do
    End If
    
    rowDiff1 = pMatched(idxMatched).row1 - IIf(idxMatched = 0, 0, pMatched(idxMatched - 1).row1)
    rowDiff2 = pMatched(idxMatched).row2 - IIf(idxMatched = 0, 0, pMatched(idxMatched - 1).row2)
    
    If rowDiff1 > 1 Or rowDiff2 > 1 Then
      '// �V�[�g�Q�ɍs�ǉ�
      If (rowDiff2 = 1) Or ((rowDiff1 > rowDiff2) And rowDiff1 <> 1) Then
        For cnt = 1 To (pMatched(idxMatched).row1 + pad1) - (pMatched(idxMatched).row2 + pad2)
          Call wkSheet2.Rows(pMatched(idxMatched).row2 + pad2 + cnt - 1).Insert(Shift:=xlDown)
          wkSheet2.Rows(pMatched(idxMatched).row2 + pad2 + cnt - 1).Interior.ColorIndex = CLR_DIFF_DEL_ROW
          Call wkSheet1.Rows(pMatched(idxMatched).row2 + pad2 + cnt - 1).Copy
          Call wkSheet2.Rows(pMatched(idxMatched).row2 + pad2 + cnt - 1).PasteSpecial(Paste:=xlValues)
          If ROW_DIFF_STRIKETHROUGH Then
            wkSheet2.Rows(pMatched(idxMatched).row2 + pad2 + cnt - 1).Font.Strikethrough = True
          End If
          Call wkSheet2.Cells(pMatched(idxMatched).row2 + pad2 + cnt - 1, 1).NoteText(FLG_DEL_ROW) '// ��ō폜
        Next
        pad2 = pad2 + (pMatched(idxMatched).row1 + pad1) - (pMatched(idxMatched).row2 + pad2)
      ElseIf (rowDiff1 = 1) Or (rowDiff1 < rowDiff2) Then
        For cnt = 1 To (pMatched(idxMatched).row2 + pad2) - (pMatched(idxMatched).row1 + pad1)
          Call wkSheet1.Rows(pMatched(idxMatched).row1 + pad1 + cnt - 1).Insert(Shift:=xlDown)
          wkSheet1.Rows(pMatched(idxMatched).row1 + pad1 + cnt - 1).Interior.ColorIndex = CLR_DIFF_DEL_ROW
          Call wkSheet2.Rows(pMatched(idxMatched).row1 + pad1 + cnt - 1).Copy
          Call wkSheet1.Rows(pMatched(idxMatched).row1 + pad1 + cnt - 1).PasteSpecial(Paste:=xlValues)
          If ROW_DIFF_STRIKETHROUGH Then
            wkSheet1.Rows(pMatched(idxMatched).row1 + pad1 + cnt - 1).Font.Strikethrough = True
          End If
          Call wkSheet2.Cells(pMatched(idxMatched).row1 + pad1 + cnt - 1, 1).NoteText(FLG_INS_ROW)  '// ��ō폜
          wkSheet2.Rows(pMatched(idxMatched).row1 + pad1 + cnt - 1).Interior.ColorIndex = CLR_DIFF_INS_ROW
        Next
        pad1 = pad1 + (pMatched(idxMatched).row2 + pad2) - (pMatched(idxMatched).row1 + pad1)
      End If
    End If
    idxMatched = idxMatched + 1
    
    If wkSheet1.UsedRange.Rows.Count > rowCntLimit Or wkSheet2.UsedRange.Rows.Count > rowCntLimit Then
      Exit Do
    End If
  Loop
  
  '// �Ō�̍s�ɂ��ĕ␳
  rowDiff = (wkSheet1.UsedRange.Row + wkSheet1.UsedRange.Rows.Count) - (wkSheet2.UsedRange.Row + wkSheet2.UsedRange.Rows.Count)
  If rowDiff > 0 Then
    For cnt = wkSheet1.UsedRange.Row + wkSheet1.UsedRange.Rows.Count - rowDiff To wkSheet1.UsedRange.Row + wkSheet1.UsedRange.Rows.Count - 1
      wkSheet2.Rows(cnt).Interior.ColorIndex = CLR_DIFF_DEL_ROW
      wkSheet1.Rows(cnt).Interior.ColorIndex = CLR_DIFF_INS_ROW
      Call wkSheet1.Rows(cnt).Copy
      Call wkSheet2.Rows(cnt).PasteSpecial(Paste:=xlValues)
      If ROW_DIFF_STRIKETHROUGH Then
        wkSheet2.Rows(cnt).Font.Strikethrough = True
      End If
      Call wkSheet2.Cells(cnt, 1).NoteText(FLG_DEL_ROW) '// ��ō폜
    Next
  ElseIf rowDiff < 0 Then
    For cnt = wkSheet2.UsedRange.Row + wkSheet2.UsedRange.Rows.Count + rowDiff To wkSheet2.UsedRange.Row + wkSheet2.UsedRange.Rows.Count - 1
      wkSheet2.Rows(cnt).Interior.ColorIndex = CLR_DIFF_INS_ROW
      wkSheet1.Rows(cnt).Interior.ColorIndex = CLR_DIFF_DEL_ROW
      Call wkSheet2.Rows(cnt).Copy
      Call wkSheet1.Rows(cnt).PasteSpecial(Paste:=xlValues)
      If ROW_DIFF_STRIKETHROUGH Then
        wkSheet1.Rows(cnt).Font.Strikethrough = True
      End If
      Call wkSheet2.Cells(cnt, 1).NoteText(FLG_INS_ROW) '// ��ō폜
    Next
      
  End If
End Sub

'// ////////////////////////////////////////////////////////////////////////////
'// END
'// ////////////////////////////////////////////////////////////////////////////

