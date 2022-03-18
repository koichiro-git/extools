VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSearch 
   Caption         =   "�g������"
   ClientHeight    =   6375
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6210
   OleObjectBlob   =   "frmSearch.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'// ////////////////////////////////////////////////////////////////////////////
'// �v���W�F�N�g   : �g���c�[��
'// �^�C�g��       : �g�������t�H�[��
'// ���W���[��     : frmSearch
'// ����           : ���K�\���ł̌������s��
'// ////////////////////////////////////////////////////////////////////////////
'// Copyright (c) by Koichiro.
'// ////////////////////////////////////////////////////////////////////////////
Option Explicit
Option Base 0


'// ////////////////////////////////////////////////////////////////////////////
'// �v���C�x�[�g�ϐ�
'// �������ʊi�[�^�C�v
Private Type udMatched
    FileName    As String
    SheetName   As String
    Row         As Long
    Col         As Integer
    TargetText  As String
    NoteText    As String
    SavedFile   As Boolean
End Type

Private pMatched()  As udMatched    '// �������ʊi�[�p�z��
Private pMatchCnt   As Long         '// �������ʐ�


'// //////////////////////////////////////////////////////////////////
'// �C�x���g�F �t�H�[����������
Private Sub UserForm_Initialize()
    '// ������̌����̓f�t�H���g��ON
    ckbSearchText.Value = True
    
    '// �R���{�{�b�N�X�ݒ�
    Call gsSetCombo(cmbTarget, CMB_SRC_TARGET, 0)
    Call gsSetCombo(cmbOutput, CMB_SRC_OUTPUT, 0)
    
    '// �L���v�V�����ݒ�
    frmSearch.Caption = LBL_SRC_FORM
    cmdDir.Caption = LBL_COM_BROWSE
    ckbSubDir.Caption = LBL_SRC_SUB_DIR
    ckbCaseSensitive.Caption = LBL_SRC_IGNORE_CASE
    fraOptions.Caption = LBL_SRC_OBJECT
    ckbSearchText.Caption = LBL_SRC_CELL_TEXT
    ckbSearchFormula.Caption = LBL_SRC_CELL_FORMULA
    ckbSearchShape.Caption = LBL_SRC_SHAPE
    ckbSearchComment.Caption = LBL_SRC_COMMENT
    ckbSearchName.Caption = LBL_SRC_CELL_NAME
    ckbSearchSheetName.Caption = LBL_SRC_SHEET_NAME
    ckbSearchLink.Caption = LBL_SRC_HYPERLINK
    ckbSearchHeader.Caption = LBL_SRC_HEADER
    ckbSearchGraph.Caption = LBL_SRC_GRAPH
    lblString.Caption = LBL_SRC_STRING
    lblTarget.Caption = LBL_SRC_TARGET
    lblMarker.Caption = LBL_SRC_MARK
    lblDir.Caption = LBL_SRC_DIR
    cmdSelectAll.Caption = LBL_COM_CHECK_ALL
    cmdClear.Caption = LBL_COM_UNCHECK
    cmdExecute.Caption = LBL_COM_EXEC
    cmdClose.Caption = LBL_COM_CLOSE
End Sub


'// //////////////////////////////////////////////////////////////////
'// �C�x���g�F ����{�^�� �N���b�N��
Private Sub cmdClose_Click()
    Call Me.Hide
End Sub


'// //////////////////////////////////////////////////////////////////
'// �C�x���g�F �Q�ƃ{�^�� �N���b�N��
Private Sub cmdDir_Click()
    Dim FilePath  As String
    
    If Not gfShowSelectFolder(0, FilePath) Then
        Exit Sub
    Else
        txtDirectory.Text = FilePath
    End If
End Sub


'// //////////////////////////////////////////////////////////////////
'// �C�x���g�F �����ΏۃR���{ �ύX��
Private Sub cmbTarget_Change()
    Select Case cmbTarget.Value
        Case 0  '// ���݂̃V�[�g
            cmdDir.Enabled = False
            ckbSubDir.Enabled = False
            txtDirectory.Enabled = False
            txtDirectory.BackColor = CLR_DISABLED
            ckbSearchSheetName.Enabled = False
            cmbOutput.Enabled = True
        Case 1  '// �u�b�N�S��
            cmdDir.Enabled = False
            ckbSubDir.Enabled = False
            txtDirectory.Enabled = False
            txtDirectory.BackColor = CLR_DISABLED
            ckbSearchSheetName.Enabled = True
            cmbOutput.Enabled = True
        Case 2  '// �f�B���N�g���P��
            cmdDir.Enabled = True
            ckbSubDir.Enabled = True
            txtDirectory.Enabled = True
            txtDirectory.BackColor = CLR_ENABLED
            ckbSearchSheetName.Enabled = True
            cmbOutput.Enabled = False
    End Select
End Sub


'// //////////////////////////////////////////////////////////////////
'// �C�x���g�F ���s�{�^�� �N���b�N��
Private Sub cmdExecute_Click()
    Dim wkSheet   As Worksheet
    Dim fs        As Object
  
    '// null�`�F�b�N
    If Trim(txtSearch.Value) = BLANK Then
      Call MsgBox(MSG_NO_CONDITION, vbOKOnly, APP_TITLE)
      Call txtSearch.SetFocus
      Exit Sub
    End If
  
    Application.ScreenUpdating = False
    
    '// �������ʃN���A
    pMatchCnt = 0
    Erase pMatched
    
    '// �������s�ipsExecSearch�Ăяo���j
    Select Case cmbTarget.Value
        Case 0  '// ���݂̃V�[�g
            Call psExecSearch(ActiveSheet, txtSearch.Text, ckbCaseSensitive.Value)
        Case 1  '// �u�b�N�S��
            For Each wkSheet In ActiveWorkbook.Sheets
                Call psExecSearch(wkSheet, txtSearch.Text, ckbCaseSensitive.Value)
            Next
        Case 2  '// �f�B���N�g���P��
            If Trim(txtDirectory.Text) <> BLANK Then
                Set fs = CreateObject("Scripting.FileSystemObject")
                
                '// �����p�X�m�F
                If fs.FolderExists(txtDirectory.Text) Then
                    Call psGetExcelFiles(fs, txtDirectory.Text, txtSearch.Text, ckbCaseSensitive.Value, ckbSubDir.Value)
                Else
                    Call MsgBox(MSG_DIR_NOT_EXIST, vbOKOnly, APP_TITLE)
                    Exit Sub
                End If
                Set fs = Nothing
            Else
                Call MsgBox(MSG_NO_DIR, vbOKOnly, APP_TITLE)
                Call txtDirectory.SetFocus
                Application.ScreenUpdating = True
                Exit Sub
            End If
    End Select
    
    If pMatchCnt > 0 Then   '// �������ʂ�1���ȏ゠��΃V�[�g�ɏo�͂��A��������
        Call psShowResult
        Call MsgBox(MSG_FINISHED, vbOKOnly, APP_TITLE)
        Call Me.Hide
    Else
        Call MsgBox(MSG_NO_RESULT, vbOKOnly, APP_TITLE)
    End If
  
    Application.StatusBar = False
    Application.ScreenUpdating = True
End Sub


'// //////////////////////////////////////////////////////////////////
'// �C�x���g�F �S�Ă�I���{�^�� �N���b�N��
Private Sub cmdSelectAll_Click()
    Call psSetCheckBoxes(True)
End Sub


'// //////////////////////////////////////////////////////////////////
'// �C�x���g�F �I�������{�^�� �N���b�N��
Private Sub cmdClear_Click()
    Call psSetCheckBoxes(False)
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// ���\�b�h�F   �����Ώۃ`�F�b�N�{�b�N�X�ݒ�
'// �����F       �����Ώۃ`�F�b�N�{�b�N�X�̒l�������̐^�U�l�Ɉꊇ�ݒ肷��B
'// �����F       newValue: �`�F�b�N�{�b�N�X�̐ݒ�l
'// �߂�l�F     �Ȃ�
'// ////////////////////////////////////////////////////////////////////////////
Private Sub psSetCheckBoxes(newValue As Boolean)
    ckbSearchText.Value = newValue
    ckbSearchFormula.Value = newValue
    ckbSearchShape.Value = newValue
    ckbSearchComment.Value = newValue
    ckbSearchName.Value = newValue
    ckbSearchSheetName.Value = newValue
    ckbSearchLink.Value = newValue
    ckbSearchHeader.Value = newValue
    ckbSearchGraph.Value = newValue
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// ���\�b�h�F   �f�B���N�g�����u�b�N����
'// �����F       �w�肳�ꂽ�f�B���N�g�����̃u�b�N����������
'// �����F       fs: �t�@�C���V�X�e���I�u�W�F�N�g
'//              dirName: �����Ώۃf�B���N�g��
'//              patternStr: ����������
'//              caseSensitive: �啶���������̋�ʃt���O
'//              searchSubDir: �T�u�f�B���N�g�������t���O
'// �߂�l�F     �Ȃ�
'// ////////////////////////////////////////////////////////////////////////////
Private Sub psGetExcelFiles(fs As Object, dirName As String, patternStr As String, caseSensitive As Boolean, searchSubDir As Boolean)
    Dim parentDir   As Object
    Dim children    As Object
    Dim wkBook      As Workbook
    Dim wkSheet     As Worksheet
    Dim isDuplName  As Boolean    '// �ΏۂƂȂ�u�b�N���J����Ă���ꍇTrue
    
    Set parentDir = fs.GetFolder(dirName)
    
    '// �t�@�C���̌���
    For Each children In parentDir.files
        With children
            If LCase(Right(.Name, 3)) = "xls" Then
                '// ����
                '// �u�b�N�����ɊJ����Ă��邩���m�F
                isDuplName = False
                For Each wkBook In Workbooks
                    If wkBook.Name = children.Name Then
                        isDuplName = True
                        Exit For
                    End If
                Next
                
                If Not isDuplName Then  '// �u�b�N���J����Ă���ꍇ�͌����ΏۊO
                    Set wkBook = Workbooks.Open(children.Path, ReadOnly:=True, password:=EXCEL_PASSWORD)
                    For Each wkSheet In wkBook.Worksheets
                        Call psExecSearch(wkSheet, patternStr, caseSensitive)
                    Next
                    Call wkBook.Close(SaveChanges:=False)
                    Set wkBook = Nothing
                End If
            End If
        End With
    Next
    
    '// �T�u�t�H���_������ꍇ�A����
    If searchSubDir Then
        For Each children In parentDir.SubFolders
          '// �q�f�B���N�g���̍ċA�Ăяo��
          Call psGetExcelFiles(fs, children.Path, patternStr, caseSensitive, True)
        Next
    End If
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// ���\�b�h�F   �f�B���N�g�����u�b�N����
'// �����F       �w�肳�ꂽ�f�B���N�g�����̃u�b�N����������
'// �����F       wkSheet: �����ΏۃV�[�g
'//              patternStr: ����������
'//              caseSensitive: �啶���������̋�ʃt���O
'// �߂�l�F     �Ȃ�
'// ////////////////////////////////////////////////////////////////////////////
Private Sub psExecSearch(wkSheet As Worksheet, patternStr As String, caseSensitive As Boolean)
    Dim regExp        As Object         '// ���K�\���I�u�W�F�N�g
    Dim targetCell    As Range
    Dim hLink         As Hyperlink
    Dim rangeName     As Name
    Dim shapeObj      As Shape
    Dim commentObj    As Comment
    Dim chartObj      As Chart
    Dim seriesObj     As Series
    Dim bffText       As String
    Dim idxChart      As Long
    Dim idxCellSrch   As Long           '// �����Z�����J�E���^
    Dim numCellCnt    As Long           '// �����ΏۃZ����
  
    numCellCnt = numCellCnt + IIf(ckbSearchText.Value, wkSheet.UsedRange.Count, 0)
    If pfGetCellCount(wkSheet.UsedRange, xlCellTypeFormulas) > -1 Then
        numCellCnt = numCellCnt + IIf(ckbSearchFormula.Value, wkSheet.UsedRange.SpecialCells(xlCellTypeFormulas).Count, 0)
    End If
  
    '// ���K�\���I�u�W�F�N�g�̍쐬
    Set regExp = CreateObject("VBScript.RegExp")
    regExp.Pattern = patternStr
    regExp.IgnoreCase = caseSensitive
  
    '// �������s�i���K�\���̋L�ڃ`�F�b�N�j
    If Not pfCheckRegExp(regExp) Then
        Call MsgBox(MSG_WRONG_COND, vbOKOnly, APP_TITLE)
        Set regExp = Nothing
        Exit Sub
    End If
  
    '// �Z������������� //////////
    If ckbSearchText.Value Then
        For Each targetCell In wkSheet.UsedRange
            If regExp.test(targetCell.Text) Then
                Call psSetMatchedRec(wkSheet, targetCell.Row, targetCell.Column, targetCell.Text, BLANK)
                
                '// �Z�����F�Ȃ�
                Select Case cmbOutput.Value
                    Case 0  '// �������Ȃ�
                    Case 1  '// �����𒅐F
                      targetCell.Font.ColorIndex = COLOR_DIFF_CELL
                    Case 2  '// �Z���𒅐F
                      targetCell.Interior.ColorIndex = COLOR_DIFF_CELL
                    Case 3  '// �g�𒅐F
                      targetCell.Borders.LineStyle = xlContinuous
                      targetCell.Borders.ColorIndex = COLOR_DIFF_CELL
                    Case 4  '// �Y���Z�����܂ލs�ȊO���\��
                      '// �����@�\
                End Select
            End If
            
            idxCellSrch = idxCellSrch + 1
            If idxCellSrch Mod 1000 = 0 Then
                Application.StatusBar = "������... [ " & wkSheet.Name & " " & CStr(CInt(idxCellSrch / numCellCnt)) & " ]"
            End If
        Next
    End If
    
    '// �������� //////////
    If ckbSearchFormula.Value And pfGetCellCount(wkSheet.UsedRange, xlCellTypeFormulas) > -1 Then
        For Each targetCell In wkSheet.UsedRange.SpecialCells(xlCellTypeFormulas)
            If regExp.test(targetCell.FormulaLocal) Then
                Call psSetMatchedRec(wkSheet, targetCell.Row, targetCell.Column, targetCell.FormulaLocal, "����")
                
                '// �Z�����F�Ȃ�
                Select Case cmbOutput.Value
                  Case 0  '// �������Ȃ�
                  Case 1  '// �����𒅐F
                    targetCell.Font.ColorIndex = COLOR_DIFF_CELL
                  Case 2  '// �Z���𒅐F
                    targetCell.Interior.ColorIndex = COLOR_DIFF_CELL
                  Case 3  '// �g�𒅐F
                    targetCell.Borders.LineStyle = xlContinuous
                    targetCell.Borders.ColorIndex = COLOR_DIFF_CELL
                  Case 4  '// �Y���Z�����܂ލs�ȊO���\��
                End Select
            End If
            
            idxCellSrch = idxCellSrch + 1
            If idxCellSrch Mod 1000 = 0 Then
                Application.StatusBar = "������... [ " & wkSheet.Name & " " & CStr(CInt(idxCellSrch / numCellCnt)) & " ]"
            End If
        Next
    End If
  
    '// �V�F�C�v���̕���������� //////////
    If ckbSearchShape.Value Then
        For Each shapeObj In wkSheet.Shapes
            If shapeObj.Type <> msoComment Then '// �V�F�C�v�̂����R�����g�ɂ��Ă̓R�����g���̂��������邽�ߏ��O
                Call psExecSearch_Shape(regExp, wkSheet, shapeObj, False)
            End If
        Next
    End If
  
    '// �R�����g���̕���������� //////////
    If ckbSearchComment.Value Then
        For Each commentObj In wkSheet.Comments
            If regExp.test(commentObj.Text) Then
                Call psSetMatchedRec(wkSheet, commentObj.Parent.Cells.Row, commentObj.Parent.Cells.Column, commentObj.Text, "�R�����g")
            End If
        Next
    End If
  
    '// �Z�����̂����� //////////
    '// ������Name������ꍇ�̃G���[��������邽�߁A���胍�W�b�N���O�����ipfCheckRangeName�j
    If ckbSearchName.Value Then
        For Each rangeName In wkSheet.Parent.Names  '// �u�b�N��Names�v���p�e�B���Q�Ƃ���K�v������i�����s���j
            If pfCheckRangeName(rangeName, wkSheet) Then
                If regExp.test(rangeName.Name) Then
                    Call psSetMatchedRec(wkSheet, rangeName.RefersToRange.Row, rangeName.RefersToRange.Column, rangeName.Name, "�Z������")
                End If
            End If
        Next
    End If
  
    '// �n�C�p�[�����N������� //////////
    If ckbSearchLink.Value Then
        For Each hLink In wkSheet.Hyperlinks
            If regExp.test(hLink.Address) Or regExp.test(hLink.SubAddress) Then
                Call psSetMatchedRec(wkSheet, hLink.Range.Row, hLink.Range.Column, hLink.Address & "[" & hLink.SubAddress & "]", "�n�C�p�[�����N")
            End If
        Next
    End If
  
  '// �V�[�g�������� //////////
    If ckbSearchSheetName.Value Then
        If regExp.test(wkSheet.Name) Then
            Call psSetMatchedRec(wkSheet, 1, 1, wkSheet.Name, "�V�[�g��")
        End If
    End If
  
  
    '// �w�b�_�ƃt�b�^�̕���������� //////////
    If ckbSearchHeader.Value Then
        If regExp.test(wkSheet.PageSetup.LeftHeader) Then
            Call psSetMatchedRec(wkSheet, 1, 1, wkSheet.PageSetup.LeftHeader, "�w�b�_�i���j")
        End If
        If regExp.test(wkSheet.PageSetup.CenterHeader) Then
            Call psSetMatchedRec(wkSheet, 1, 1, wkSheet.PageSetup.CenterHeader, "�w�b�_�i�����j")
        End If
        If regExp.test(wkSheet.PageSetup.RightHeader) Then
            Call psSetMatchedRec(wkSheet, 1, 1, wkSheet.PageSetup.RightHeader, "�w�b�_�i�E�j")
        End If
        If regExp.test(wkSheet.PageSetup.LeftFooter) Then
            Call psSetMatchedRec(wkSheet, 1, 1, wkSheet.PageSetup.LeftFooter, "�t�b�^�i���j")
        End If
        If regExp.test(wkSheet.PageSetup.CenterFooter) Then
            Call psSetMatchedRec(wkSheet, 1, 1, wkSheet.PageSetup.CenterFooter, "�t�b�^�i�����j")
        End If
        If regExp.test(wkSheet.PageSetup.RightFooter) Then
            Call psSetMatchedRec(wkSheet, 1, 1, wkSheet.PageSetup.RightFooter, "�t�b�^�i�E�j")
        End If
    End If
  
    '// �O���t������ //////////
    If ckbSearchGraph.Value Then
        For idxChart = 1 To wkSheet.ChartObjects.Count  '// �`���[�g�̔z��͂P����J�n
            Set chartObj = wkSheet.ChartObjects(idxChart).Chart
            If regExp.test(pfGetChartTitle(chartObj)) Then
                Call psSetMatchedRec(wkSheet, 1, 1, chartObj.ChartTitle.Characters.Text, "�`���[�g�^�C�g��")
            End If
            
            For Each seriesObj In chartObj.SeriesCollection
                If regExp.test(seriesObj.Name) Then
                    Call psSetMatchedRec(wkSheet, 1, 1, seriesObj.Name, "�`���[�g�n��")
                End If
            Next
        Next
    End If
    
    Set regExp = Nothing
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// ���\�b�h�F   �V�F�C�v���e�L�X�g�擾
'// �����F       �V�F�C�v���̃e�L�X�g���擾����BCharacters���\�b�h���T�|�[�g���Ȃ��ꍇ�͗�O�����Ńn���h�����O
'//              psExecSearch_Shape�œ��肳�ꂽ�V�F�C�v���̃e�L�X�g��߂�
'// �����F       shapeObj: �ΏۃV�F�C�v�I�u�W�F�N�g
'// �߂�l�F     �V�F�C�v���̃e�L�X�g�B�V�F�C�v���e�L�X�g���T�|�[�g���Ă��Ȃ��ꍇ�͈ꗥ�Ńu�����N
'// ////////////////////////////////////////////////////////////////////////////
Private Function pfGetShapeText(shapeObj As Shape) As String
On Error GoTo ErrorHandler
    If shapeObj.Type = msoTextEffect Then '// ���[�h�A�[�g�̏ꍇ
        pfGetShapeText = shapeObj.TextEffect.Text
    Else
        pfGetShapeText = shapeObj.TextFrame.Characters.Text
    End If
Exit Function

ErrorHandler:
    pfGetShapeText = BLANK
End Function


'// ////////////////////////////////////////////////////////////////////////////
'// ���\�b�h�F   �����F�V�F�C�v
'// �����F       �V�F�C�v���̕��������������B�O���[�v������Ă���ꍇ�͍ċA�������s���B
'// �����F       regExp: ���K�\���I�u�W�F�N�g
'//              wkSheet: �ΏۃV�[�g
'//              shapeObj: �ΏۃV�F�C�v�I�u�W�F�N�g
'//              isGrouped: �O���[�v���I�u�W�F�N�g���ۂ��i�ċA�Ăяo������Ă��邩�j
'// �߂�l�F     �Ȃ�
'// ////////////////////////////////////////////////////////////////////////////
Private Sub psExecSearch_Shape(regExp As Object, wkSheet As Worksheet, shapeObj As Shape, isGrouped As Boolean)
    Dim bffText   As String
    Dim subShape  As Shape
    
    If shapeObj.Type = msoGroup Then
        For Each subShape In shapeObj.GroupItems
            Call psExecSearch_Shape(regExp, wkSheet, subShape, True)
        Next
    Else
        bffText = pfGetShapeText(shapeObj)
        If bffText <> BLANK Then
            If regExp.test(bffText) Then
                Call psSetMatchedRec(wkSheet, IIf(isGrouped, -1, shapeObj.TopLeftCell.Row), IIf(isGrouped, -1, shapeObj.TopLeftCell.Column), bffText, "�V�F�C�v�F" & shapeObj.Name)
            End If
        End If
    End If
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// ���\�b�h�F   �������ʏo��
'// �����F       �������ʂ�ʃu�b�N�ŏo�͂���
'// �����F       �Ȃ�
'// �߂�l�F     �Ȃ�
'// ////////////////////////////////////////////////////////////////////////////
Private Sub psShowResult()
    Dim wkSheet   As Worksheet
    Dim idxRow    As Long
    
    '// �o�͐�̐ݒ�
    With Workbooks.Add
        Set wkSheet = .ActiveSheet
    End With
  
    '// �w�b�_�̐ݒ�
    Call gsDrawResultHeader(wkSheet, HDR_SEARCH, 1)
  
    wkSheet.Cells.NumberFormat = "@"
    
    '// �l�̐ݒ�
    For idxRow = 0 To UBound(pMatched) - 1
        wkSheet.Cells(idxRow + 2, 1).Value = pMatched(idxRow).FileName
        wkSheet.Cells(idxRow + 2, 2).Value = pMatched(idxRow).SheetName
        wkSheet.Cells(idxRow + 2, 3).Value = IIf(pMatched(idxRow).Row > 0, mdlCommon.gfGetColIndexString(pMatched(idxRow).Col) & CStr(pMatched(idxRow).Row), BLANK)
        wkSheet.Cells(idxRow + 2, 4).Value = pMatched(idxRow).TargetText
        wkSheet.Cells(idxRow + 2, 5).Value = pMatched(idxRow).NoteText
        
        If pMatched(idxRow).SavedFile And pMatched(idxRow).Row > 0 Then '// �Z�[�u����Ă���Ƃ��̂݃����N�ݒ�
            ActiveSheet.Hyperlinks.Add Anchor:=wkSheet.Cells(idxRow + 2, 3), Address:=wkSheet.Cells(idxRow + 2, 1).Value, SubAddress:="'" & wkSheet.Cells(idxRow + 2, 2).Value & "'!" & wkSheet.Cells(idxRow + 2, 3).Value
        End If
    Next
  
    '// //////////////////////////////////////////////////////
    '// �����̐ݒ�
    '// ���̐ݒ�
    wkSheet.Columns("A:C").ColumnWidth = 10
    wkSheet.Columns("D:E").ColumnWidth = 30
    
    '// �g���̐ݒ�
    Call gsPageSetup_Lines(wkSheet, 1)
    
'    Call wkSheet.Range(wkSheet.Cells(1, 1), wkSheet.Cells(UBound(pMatched) + 1, 5)).Select
'    Call mdlCommon.gsDrawLine_Data
'
'    '// �w�b�_�̏C��
'    Call wkSheet.Range("A1:E1").Select
'    Call mdlCommon.gsDrawLine_Header
    
    '//�t�H���g
    wkSheet.Cells.Font.Name = APP_FONT
    wkSheet.Cells.Font.Size = APP_FONT_SIZE
    
    Call wkSheet.Cells(1, 1).Select
    
    '// �㏈��
    Call wkSheet.Cells(1, 1).Select
    ActiveWorkbook.Saved = True
    Application.StatusBar = False
    Application.ScreenUpdating = True
    '// ����Ƃ��ɕۑ������߂Ȃ�
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// ���\�b�h�F   �����q�b�g���R�[�h�o�^
'// �����F       �����Ƀq�b�g�������e��z��ɓo�^����
'// �����F       wkSheet: �Ώۃ��[�N�V�[�g
'//              Row: �q�b�g�����s
'//              Col: �q�b�g������
'//              TargetText: �q�b�g�����l
'//              NoteText: ���l
'// �߂�l�F     �Ȃ�
'// ////////////////////////////////////////////////////////////////////////////
Private Sub psSetMatchedRec(wkSheet As Worksheet, Row As Long, Col As Integer, TargetText As String, NoteText As String)
    ReDim Preserve pMatched(pMatchCnt + 1)
    
    pMatched(pMatchCnt).FileName = wkSheet.Parent.Path & "\" & wkSheet.Parent.Name
    pMatched(pMatchCnt).SheetName = wkSheet.Name
    pMatched(pMatchCnt).Row = Row
    pMatched(pMatchCnt).Col = Col
    pMatched(pMatchCnt).TargetText = TargetText
    pMatched(pMatchCnt).NoteText = NoteText
    pMatched(pMatchCnt).SavedFile = IIf(wkSheet.Parent.Path = BLANK, False, True)
    
    pMatchCnt = pMatchCnt + 1
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// ���\�b�h�F   �Z���͈̓J�E���g�擾
'// �����F       SpecialCells �̌��ʃJ�E���g�����擾����
'// �����F       targetRange: �Ώ۔͈�
'//              cellType: �擾�^�C�v
'// �߂�l�F     �͈͓��̑ΏۃZ�����B�Z�����[���̏ꍇ�� -1 ��Ԃ�
'// ////////////////////////////////////////////////////////////////////////////
Private Function pfGetCellCount(targetRange As Range, cellType As Long) As Double
On Error GoTo ErrorHandler
    pfGetCellCount = targetRange.SpecialCells(cellType).Count
    Exit Function

ErrorHandler:
    pfGetCellCount = -1
End Function


'// ////////////////////////////////////////////////////////////////////////////
'// ���\�b�h�F   ����������̑Ó�������
'// �����F       �w�肳�ꂽ���������񂪐��K�\���Ƃ��đÓ����i�G���[���������Ȃ����j���m�F����
'// �����F       regExp: ���K�\���I�u�W�F�N�g
'// �߂�l�F     �����̐���
'// ////////////////////////////////////////////////////////////////////////////
Private Function pfCheckRegExp(regExp As Object) As Boolean
On Error GoTo ErrorHandler
    pfCheckRegExp = regExp.test(BLANK)
    pfCheckRegExp = True
    Exit Function

ErrorHandler:
    pfCheckRegExp = False
End Function


'// ////////////////////////////////////////////////////////////////////////////
'// ���\�b�h�F   �Z�����̂̑Ó�������
'// �����F       �w�肳�ꂽ�Z�����̂�wkSheet�Ɋ܂܂�Ă��邩�A����їL���Ȗ��̂ł��邩�𔻒肷��
'// �����F       rangeName: �ΏۂƂȂ�Z�����̃I�u�W�F�N�g
'//              wkSheet: �ΏۂƂȂ�V�[�g
'// �߂�l�F     �Ó����̐���
'// ////////////////////////////////////////////////////////////////////////////
Private Function pfCheckRangeName(rangeName As Name, wkSheet As Worksheet) As Boolean
On Error GoTo ErrorHandler
    pfCheckRangeName = (rangeName.RefersToRange.Worksheet.Name = wkSheet.Name)
    Exit Function

ErrorHandler:
    pfCheckRangeName = False
End Function


'// ////////////////////////////////////////////////////////////////////////////
'// ���\�b�h�F   �`���[�g�^�C�g���擾
'// �����F       �w�肳�ꂽ�`���[�g�^�C�g����characters��Ԃ��B
'// �����F       chartObj: �ΏۂƂȂ�`���[�g�I�u�W�F�N�g
'// �߂�l�F     �`���[�g�̃^�C�g��������B�擾�s�̏ꍇ�͋󔒕�����
'// ////////////////////////////////////////////////////////////////////////////
Private Function pfGetChartTitle(chartObj As Chart) As String
On Error GoTo ErrorHandler
    pfGetChartTitle = chartObj.ChartTitle.Characters.Text
    Exit Function

ErrorHandler:
    pfGetChartTitle = BLANK
End Function


'// ////////////////////////////////////////////////////////////////////////////
'// END
'// ////////////////////////////////////////////////////////////////////////////