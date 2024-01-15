VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmFileList 
   Caption         =   "�t�@�C���ꗗ�o��"
   ClientHeight    =   4455
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6150
   OleObjectBlob   =   "frmFileList.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "frmFileList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'// ////////////////////////////////////////////////////////////////////////////
'// �v���W�F�N�g   : �g���c�[��
'// �^�C�g��       : �t�@�C���ꗗ�o�̓t�H�[��
'// ���W���[��     : frmFileList
'// ����           : �V�[�g�̔�r���s��
'// ////////////////////////////////////////////////////////////////////////////
'// Copyright (c) by Koichiro.
'// ////////////////////////////////////////////////////////////////////////////
Option Explicit
Option Base 0

'// ////////////////////////////////////////////////////////////////////////////
'// �v���C�x�[�g�萔
Private Const pUNLIMITED_DEPTH    As Integer = "32767"

'// �v���C�x�[�g�ϐ�
Private pRootDir        As String   '// �f�B���N�g���擾�J�n�ʒu
Private pExtentions()   As String   '// �g���q�̔z��
Private pMaxDepth       As Integer  '// �ő�[�x


'// //////////////////////////////////////////////////////////////////
'// �C�x���g�F �t�H�[�� ��������
Private Sub UserForm_Initialize()
    '// �R���{�{�b�N�X�ݒ�
    Call gsSetCombo(cmbDirDepth, CMB_LST_DEPTH, 9)
    Call gsSetCombo(cmbTargetFile, CMB_LST_TARGET, 0)
    Call gsSetCombo(cmbFileSize, CMB_LST_SIZE, 0)
    
    '// �L���v�V�����ݒ�
    frmFileList.Caption = LBL_LST_FORM
    cmdExecute.Caption = LBL_COM_EXEC
    cmdClose.Caption = LBL_COM_CLOSE
    cmdRootDir.Caption = LBL_COM_BROWSE
    ckbPath.Caption = LBL_LST_REL_PATH
    ckbHyperLink.Caption = LBL_COM_HYPERLINK
    lblRoot.Caption = LBL_LST_ROOT
    lblDepth.Caption = LBL_LST_DEPTH
    lblTarget.Caption = LBL_LST_TARGET
    lblExtentions.Caption = LBL_LST_EXT
    lblSize.Caption = LBL_LST_SIZE
End Sub


'// //////////////////////////////////////////////////////////////////
'// �C�x���g�F ����{�^�� �N���b�N��
Private Sub cmdClose_Click()
    Call Me.Hide
End Sub


'// //////////////////////////////////////////////////////////////////
'// �C�x���g�F �Q�ƃ{�^�� �N���b�N��
Private Sub cmdRootDir_Click()
    Dim FilePath  As String
    
    If Not gfShowSelectFolder(0, FilePath) Then
        Exit Sub
    Else
        txtRootDir.Text = FilePath
    End If
End Sub


'// //////////////////////////////////////////////////////////////////
'// �C�x���g�F �Ώۃt�@�C���R���{ �X�V��
Private Sub cmbTargetFile_Change()
    Select Case cmbTargetFile.Value
        Case "0"
            txtExtentions.Enabled = False
            txtExtentions.BackColor = CLR_DISABLED
        Case Else
            txtExtentions.Enabled = True
            txtExtentions.BackColor = CLR_ENABLED
    End Select
End Sub


'// //////////////////////////////////////////////////////////////////
'// �C�x���g�F ���s�{�^�� �N���b�N��
Private Sub cmdExecute_Click()
    If Trim(txtRootDir.Text) = BLANK Then  '//�󔒃`�F�b�N
        Call MsgBox(MSG_NO_DIR, vbOKOnly, APP_TITLE)
        Call txtRootDir.SetFocus
    Else
        Call gsSuppressAppEvents
        Call psShowFileList
        Call gsResumeAppEvents
        Call Me.Hide
    End If
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// ���\�b�h�F   ���X�g�o�̓��C��
'// �����F       ���X�g�o�͂��s���B
'// �����F       �Ȃ�
'// �߂�l�F     �Ȃ�
'// ////////////////////////////////////////////////////////////////////////////
Private Sub psShowFileList()
On Error GoTo ErrorHandler
    Dim fs          As Object
    Dim rootDir     As Object
    Dim wkSheet     As Worksheet
    Dim sizeUnit    As Double
    Dim sizeUnitTxt As String
    Dim sizeFormat  As String
    Dim idx         As Integer
  
    '// �ݒ�l�̋L��
    pRootDir = txtRootDir.Text                      '// ���[�g�̐ݒ�
    pExtentions = Split(txtExtentions.Text, ";")    '// �g���q (trim�����v)
    For idx = 0 To UBound(pExtentions)
        pExtentions(idx) = LCase(Trim(pExtentions(idx)))
    Next
    pMaxDepth = CInt(cmbDirDepth.Value)             '// �ő�[�x
    If cmbFileSize.Value = 0 Then                   '// �t�@�C���T�C�Y�P��
        sizeUnit = 1
        sizeUnitTxt = "B"
        sizeFormat = "#,##0 "
    ElseIf cmbFileSize.Value = 1 Then
        sizeUnit = 1024
        sizeUnitTxt = "KB"
        sizeFormat = "#,##0.0_ "
    ElseIf cmbFileSize.Value = 2 Then
        sizeUnit = 1048576
        sizeUnitTxt = "MB"
        sizeFormat = "#,##0.0_ "
    End If
    
    '// �t�@�C���V�X�e���I�u�W�F�N�g�̍쐬�ƌ����p�X�m�F
    Set fs = CreateObject("Scripting.FileSystemObject")
    If Not fs.FolderExists(pRootDir) Then
        Call MsgBox(MSG_DIR_NOT_EXIST, vbOKOnly, APP_TITLE)
        Set fs = Nothing
        Exit Sub
    End If
 
    Call Workbooks.Add
    Set wkSheet = ActiveWorkbook.ActiveSheet
  
    '// �w�b�_�̕`��
    Call gsDrawResultHeader(wkSheet, Replace(HDR_LST, "$", sizeUnitTxt), 1)
  
    '// ���[�g�̏o��
    Set rootDir = fs.GetFolder(pRootDir)
    wkSheet.Cells(2, 1).Value = rootDir.Path
    If Not rootDir.IsRootFolder Then
        wkSheet.Cells(2, 3).Value = rootDir.DateCreated
        wkSheet.Cells(2, 4).Value = rootDir.DateLastModified
    End If
    '// �h��Ԃ�
    Range(wkSheet.Cells(2, 1), wkSheet.Cells(2, 8)).Interior.ColorIndex = COLOR_ROW
    
    '// �t�@�C���o�̓��[�`���̌Ăяo���i�ċA�j
    Call psGetFileList(wkSheet, fs, rootDir.Path, 2, 0, 0, cmbTargetFile.Value, ckbHyperLink.Value, sizeUnit)
    Set fs = Nothing
  
    '// //////////////////////////////////////////////////////
    '// �����̐ݒ�
    '// ��̏���
    wkSheet.Columns("A:B").NumberFormatLocal = "@"              '// �p�X�A�t�@�C����
    wkSheet.Columns("C:D").NumberFormatLocal = "yyyy/mm/dd"     '// �쐬���A�X�V��
    
    wkSheet.Columns("E").NumberFormatLocal = sizeFormat         '// �t�@�C���T�C�Y
    
    '// ���̐ݒ�
    wkSheet.Columns("A").ColumnWidth = 15
    wkSheet.Columns("B").ColumnWidth = 20
    wkSheet.Columns("C:D").ColumnWidth = 9
    
    '// �g���̐ݒ�
    Call gsPageSetup_Lines(wkSheet, 1)
    
    '//�t�H���g
    wkSheet.Cells.Font.Name = APP_FONT
    wkSheet.Cells.Font.Size = APP_FONT_SIZE
    
    '// �t�@�C�����������L��
    wkSheet.Cells(1, 7).AddComment ("rhsa: Read only, Hidden, System file, Archive")
    
    '// �㏈��
    Call wkSheet.Cells(1, 1).Select
    ActiveWorkbook.Saved = True
    Exit Sub

ErrorHandler:
    Call gsShowErrorMsgDlg("frmFileList.pfShowFileList", Err, Nothing)
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// ���\�b�h�F   �t�@�C���ꗗ�o��
'// �����F       �����̃f�B���N�g���ȉ��̗v�f���o�͂���
'// �����F       wkSheet: �o�͑ΏۃV�[�g
'//              fs: �Ώۃt�@�C���V�X�e���I�u�W�F�N�g
'//              dirName: �Ώۃf�B���N�g����
'//              idxRow: ���ʏo�͍s
'//              depth: �f�B���N�g���[�x
'//              mode_Dir: �f�B���N�g���������[�h 0:�S��,1:������O,2:��̂�
'//              mode_File: �t�@�C���������[�h
'//              addLink: �n�C�p�[�����N�t��
'// �߂�l�F     �Ȃ�
'// ////////////////////////////////////////////////////////////////////////////
Private Sub psGetFileList(wkSheet As Worksheet, fs As Object, dirName As String, ByRef idxRow As Long, depth As Integer, mode_Dir As Integer, mode_File As Integer, addLink As Boolean, sizeUnit As Double)
On Error GoTo ErrorHandler
    Dim currentRow  As Long
    Dim parentDir   As Object
    Dim children    As Object
    Dim cnt         As Integer
    Dim isTarget    As Boolean
    Dim isEmptyDir  As Boolean
  
    currentRow = idxRow
  
    Set parentDir = fs.GetFolder(dirName)
    isEmptyDir = True
  
    '// �X�e�[�^�X�o�[�X�V
    Application.StatusBar = MSG_PROCESSING & " [ " & dirName & " ]"
    
    '// �t�@�C���̏o��
    For Each children In parentDir.files
        isEmptyDir = False
        With children
            Select Case mode_File
                Case "0"
                    isTarget = True
                Case "1"
                    isTarget = False
                    For cnt = 0 To UBound(pExtentions)
                        If LCase(Right(.Name, Len(pExtentions(cnt)))) = pExtentions(cnt) Then
                            isTarget = True
                            Exit For
                        End If
                    Next
                Case "2"
                    isTarget = True
                    For cnt = 0 To UBound(pExtentions)
                        If LCase(Right(.Name, Len(pExtentions(cnt)))) = pExtentions(cnt) Then
                            isTarget = False
                            Exit For
                        End If
                    Next
            End Select
            
            If isTarget Then
                idxRow = idxRow + 1
                wkSheet.Cells(idxRow, 2).Value = .Name
                wkSheet.Cells(idxRow, 3).Value = .DateCreated
                wkSheet.Cells(idxRow, 4).Value = .DateLastModified
                wkSheet.Cells(idxRow, 5).Value = .Size / sizeUnit
                wkSheet.Cells(idxRow, 6).Value = .Type
                wkSheet.Cells(idxRow, 7).Value = pfGetAttrString(.Attributes)
                
                '// �[���o�C�g�t�@�C���̔��l��
                If .Size = 0 Then
                    wkSheet.Cells(idxRow, 8).Value = MSG_ZERO_BYTE
                End If
                '// �����N�̐ݒ�
                If addLink Then
                    Call wkSheet.Cells(idxRow, 2).Hyperlinks.Add(Anchor:=wkSheet.Cells(idxRow, 2), Address:=.parentfolder & "\" & .Name)
                End If
            End If
        End With
    Next
  
    '// �T�u�t�H���_�̏o��
    For Each children In parentDir.SubFolders
        isEmptyDir = False
        idxRow = idxRow + 1
        With children
            wkSheet.Cells(idxRow, 1).Value = IIf(ckbPath.Value, "." & Mid(.Path, Len(pRootDir) + 1), .Path)
            wkSheet.Cells(idxRow, 3).Value = .DateCreated
            wkSheet.Cells(idxRow, 4).Value = .DateLastModified
            '// �h��Ԃ�
            Range(wkSheet.Cells(idxRow, 1), wkSheet.Cells(idxRow, 8)).Interior.ColorIndex = COLOR_ROW
            
            '// �����N�̐ݒ�
            If addLink Then
                Call wkSheet.Cells(idxRow, 1).Hyperlinks.Add(Anchor:=Cells(idxRow, 1), Address:=.Path)
            End If
            
            '// �q�f�B���N�g���̍ċA�Ăяo��
            If depth < pMaxDepth Then
                Call psGetFileList(wkSheet, fs, .Path, idxRow, depth + 1, mode_Dir, mode_File, addLink, sizeUnit)
            Else
                wkSheet.Cells(idxRow, 5).Value = fs.GetFolder(.Path).Size / sizeUnit    '// �z���̃t�@�C���T�C�Y���擾
                wkSheet.Cells(idxRow, 8).Value = MSG_MAX_DEPTH
            End If
        End With
    Next

    '// ��t�H���_�̔��l��
    If isEmptyDir Then
        wkSheet.Cells(currentRow, 8).Value = MSG_EMPTY_DIR
    End If
    Exit Sub
  
ErrorHandler:
    If Err.Number = 70 Then
        wkSheet.Cells(currentRow, 8).Value = MSG_ERR_PRIV
    Else
        Call gsShowErrorMsgDlg("frmFileList.psGetFileList", Err, Nothing)
    End If
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// ���\�b�h�F   �����\�������񐶐�
'// �����F       �����̑������l�𕶎���ɕϊ�����
'// �����F       targetVal: ������\�������l
'// �߂�l�F     ������\�������� "rhsa" ����
'// ////////////////////////////////////////////////////////////////////////////
Private Function pfGetAttrString(targetVal As Integer)
    pfGetAttrString = IIf(targetVal And vbReadOnly, "r", "-") & IIf(targetVal And vbHidden, "h", "-") & IIf(targetVal And vbSystem, "s", "-") & IIf(targetVal And vbArchive, "a", "-")
End Function


'// ////////////////////////////////////////////////////////////////////////////
'// END
'// ////////////////////////////////////////////////////////////////////////////
