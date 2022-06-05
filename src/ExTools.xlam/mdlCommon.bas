Attribute VB_Name = "mdlCommon"
'// ////////////////////////////////////////////////////////////////////////////
'// �v���W�F�N�g   : �g���c�[��
'// �^�C�g��       : ���ʊ֐�
'// ���W���[��     : mdlCommon
'// ����           : �V�X�e���̋��ʊ֐��A�N�����̐ݒ�Ȃǂ��Ǘ�
'// ////////////////////////////////////////////////////////////////////////////
'// Copyright (c) by Koichiro.
'// ////////////////////////////////////////////////////////////////////////////
Option Explicit
Option Base 0

'// ////////////////////////////////////////////////////////////////////////////
'// �J�X�^�}�C�Y�\�p�����[�^�i�萔�j

Public Const APP_FONT                 As String = "Meiryo UI"                                       '// #001 �\���t�H���g����
Public Const APP_FONT_SIZE            As Integer = 9                                                '// #002 �\���t�H���g�T�C�Y
Public Const HED_LEFT                 As String = ""                                                '// #003 �w�b�_������i���j
Public Const HED_CENTER               As String = ""                                                '// #004 �w�b�_������i�����j
Public Const HED_RIGHT                As String = ""                                                '// #005 �w�b�_������i�E�j
Public Const FOT_LEFT                 As String = "&""" & APP_FONT & ",�W��""&8&F / &A"             '// #006 �t�b�^������i���j
Public Const FOT_CENTER               As String = "&""" & APP_FONT & ",�W��""&8&P / &N"             '// #007 �t�b�^������i�����j
Public Const FOT_RIGHT                As String = "&""" & APP_FONT & ",�W��""&8�������: &D &T"     '// #008 �t�b�^������i�E�j
Public Const MRG_LEFT                 As Double = 0.25                                              '// #009 ����}�[�W���i���j
Public Const MRG_RIGHT                As Double = 0.25                                              '// #010 ����}�[�W���i�E�j
Public Const MRG_TOP                  As Double = 0.75                                              '// #011 ����}�[�W���i��j
Public Const MRG_BOTTOM               As Double = 0.75                                              '// #012 ����}�[�W���i���j
Public Const MRG_HEADER               As Double = 0.3                                               '// #013 ����}�[�W���i�w�b�_�j
Public Const MRG_FOOTER               As Double = 0.3                                               '// #014 ����}�[�W���i�t�b�^�j

'Public Const COLOR_HEADER             As Integer = 36                                               '// #009 ��w�b�_�F
Public Const COLOR_ROW                As Integer = 35                                               '// #018 �s�F�����F
Public Const COLOR_DIFF_CELL          As Integer = 3                                                '// �F�F3=��
Public Const COLOR_DIFF_ROW_INS       As Integer = 34                                               '// $mod
Public Const COLOR_DIFF_ROW_DEL       As Integer = 15                                               '// $mod
Public Const EXCEL_PASSWORD           As String = ""                                                '// #017 �G�N�Z�����J���ۂ̃p�X���[�h
Public Const STAT_INTERVAL            As Integer = 100                                              '// �X�e�[�^�X�o�[�X�V�p�x
Public Const ROW_DIFF_STRIKETHROUGH   As Boolean = True                                             '// $mod
Private Const MAX_COL_LEN             As Integer = 80                                               '// �N���b�v�{�[�h�ɃR�s�[����ۂ̗�ő咷
Private Const MENU_NUM                As Integer = 30                                               '// �V�[�g�����j���[�ɕ\������ۂ̃O���[�v臒l


'// ////////////////////////////////////////////////////////////////////////////
'// �A�v���P�[�V�����萔

'// �o�[�W����
Public Const APP_VERSION              As String = "2.1.1.49"                                        '// {���W���[}.{�@�\�C��}.{�o�O�C��}.{�J�����Ǘ��p}

'// �V�X�e���萔
Public Const BLANK                    As String = ""                                                '// �󔒕�����
Public Const CHR_ESC                  As Long = 27                                                  '// Escape �L�[�R�[�h
Public Const CLR_ENABLED              As Long = &H80000005                                          '// �R���g���[���w�i�F �L��
Public Const CLR_DISABLED             As Long = &H8000000F                                          '// �R���g���[���w�i�F ����
Public Const TYPE_RANGE               As String = "Range"                                           '// selection �^�C�v�F�����W
Public Const TYPE_SHAPE               As String = "Shape"                                           '// selection �^�C�v�F�V�F�C�v�ivarType�j
Public Const MENU_PREFIX              As String = "sheet"
Public Const EXCEL_FILE_EXT           As String = "*.xls; *.xlsx"                                   '// �G�N�Z���g���q


'// ////////////////////////////////////////////////////////////////////////////
'// Windows API �֘A�̐錾

'// �萔
Private Const BIF_RETURNONLYFSDIRS = &H1
Private Const MAX_PATH = 260

'// �^�C�v
Private Type BROWSEINFO
    hwndOwner       As Long
    pidlRoot        As Long
    pszDisplayName  As String
    lpszTitle       As String
    ulFlags         As Long
    lpfn            As Long
    lParam          As Long
    iImage          As Long
End Type

'// �t�H���_�I��
Private Declare Function apiSHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolder" (lpBrowseInfo As BROWSEINFO) As Long
'// �p�X�擾
Private Declare Function apiSHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDList" (ByVal piDL As Long, ByVal strPath As String) As Long
'//�L�[���荞��
Public Declare Function GetAsyncKeyState Lib "User32.dll" (ByVal vKey As Long) As Long


'// ////////////////////////////////////////////////////////////////////////////
'// �p�u���b�N�ϐ�

'// �͈̓^�C�v
Public Type udTargetRange
    minRow  As Long
    minCol  As Integer
    maxRow  As Long
    maxCol  As Integer
    Rows    As Long
    Columns As Integer
End Type

Public gADO                             As cADO         '// �ڑ���DB/Excel�I�u�W�F�N�g
Public gLang                            As Long         '// ����


'// ////////////////////////////////////////////////////////////////////////////
'// �v���C�x�[�g�ϐ�
Private pVLookupMaster                  As String               '// VLookUp�R�s�[�@�\�Ń}�X�^�\�͈͂��i�[����
Private pVLookupMasterIndex             As String               '// VLookUp�R�s�[�@�\�Ń}�X�^�\�͈͂̕\���C���f�N�X���i�[����v


'// ////////////////////////////////////////////////////////////////////////////
'// ���\�b�h�F   �c�[��������
'// �����F       ���j���[�̍\���A�A�v���I�u�W�F�N�g�̐ݒ���s���B
'// ////////////////////////////////////////////////////////////////////////////
Private Sub psInitExTools()
    '// ����̐ݒ�
    gLang = Application.LanguageSettings.LanguageID(msoLanguageIDInstall)
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// ���\�b�h�F   ���j���[�ǉ��֐�
'// �����F       ���j���[�̒ǉ����s���B�e�֐��i���j���[�\���֐��j����Ăяo�����B
'// �����F       barCtrls:      �e�o�[�R���g���[��
'//              menuCaption:   �L���v�V����
'//              actionCommand: �N���b�N���̃C�x���g�v���V�[�W��
'//              iconNum:       �A�C�R���ԍ�
'//              groupFlag:     �O���[�v���v��
'//              functionID:    �p�����[�^
'//              menuEnabled:   ���j���[�̗L��/����
'// ////////////////////////////////////////////////////////////////////////////
Private Sub psPutMenu(barCtrls As CommandBarControls, menuCaption As String, actionCommand As String, iconNum As Integer, groupFlag As Boolean, functionID As String, menuEnabled As Boolean)
    With barCtrls.Add
        .Caption = menuCaption
        .OnAction = actionCommand
        .FaceId = iconNum
        .BeginGroup = groupFlag
        .Parameter = functionID
        .Enabled = menuEnabled
    End With
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// ���\�b�h�F   �G���[���b�Z�[�W�\��
'// �����F       ��O�������ŏ����ł��Ȃ���O�̃G���[�̓��e���A�_�C�A���O�\������B
'// �����F       errSource: �G���[�̔������̃I�u�W�F�N�g�܂��̓A�v���P�[�V�����̖��O������������
'//              e: �u�a�G���[�I�u�W�F�N�g
'//              objAdo�F ADO�I�u�W�F�N�g�i�ȗ��j
'// ////////////////////////////////////////////////////////////////////////////
Public Sub gsShowErrorMsgDlg(errSource As String, ByVal e As ErrObject, Optional ado As cADO = Nothing)
    If ado Is Nothing Then
        '// ADO�I�u�W�F�N�g������̏ꍇ��VB�G���[�Ƃ��Ĉ���
        Call MsgBox(MSG_ERR & vbLf & vbLf _
                   & "Error Number: " & e.Number & vbLf _
                   & "Error Source: " & errSource & vbLf _
                   & "Error Description: " & e.Description _
                   , , APP_TITLE)
        Call e.Clear
    ElseIf ado.NativeError <> 0 Then
        '// DB�ł̃G���[�̏ꍇ
        Call MsgBox(MSG_ERR & vbLf & vbLf _
                   & "Error Number: " & ado.NativeError & vbLf _
                   & "Error Source: Database" & vbLf _
                   & "Error Description: " & ado.ErrorText _
                   , , APP_TITLE)
        ado.InitError
    ElseIf ado.ErrorCode <> 0 Then
        '// ADO�ł̃G���[�̏ꍇ
        Call MsgBox(MSG_ERR & vbLf & vbLf _
                   & "Error Number: " & ado.ErrorCode & vbLf _
                   & "Error Source: ADO" & vbLf _
                   & "Error Description�F " & ado.ErrorText _
                   , , APP_TITLE)
        ado.InitError
    Else
        '// ��L�Ŏ�蓦�����ꍇ��VB�G���[�Ƃ��Ĉ���
        Call MsgBox(MSG_ERR & vbLf & vbLf _
                   & "Error Number: " & e.Number & vbLf _
                   & "Error Source: " & errSource & vbLf _
                   & "Error Description: " & e.Description _
                   , , APP_TITLE)
        Call e.Clear
    End If
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// ���\�b�h�F   �V�[�g���ёւ�
'// �����F       �V�[�g���ŕ��ёւ���
'// �����F       sortMode: �����܂��͍~����\��������iASC/DESC�j
'// ////////////////////////////////////////////////////////////////////////////
Private Sub psSortWorksheet(sortMode As String)
    Dim i           As Integer
    Dim j           As Integer
    Dim wkSheet     As Worksheet
    Dim isOrderAsc  As Boolean
  
    '// �u�b�N���ی삳��Ă���ꍇ�ɂ̓G���[�Ƃ���
    If ActiveWorkbook.ProtectStructure Then
        Call MsgBox(MSG_BOOK_PROTECTED, vbOKOnly, APP_TITLE)
        Exit Sub
    End If
    
    '// ���s�m�F
    If MsgBox(MSG_CONFIRM, vbOKCancel, APP_TITLE) = vbCancel Then
        Exit Sub
    End If
    
    Call gsSuppressAppEvents
    
    isOrderAsc = (sortMode = "ASC") '// ����/�~���̐ݒ�
    
    '// �\�[�g
    For i = 1 To Worksheets.Count - 1
        Set wkSheet = Worksheets(i)
        
        For j = i + 1 To Worksheets.Count
            If isOrderAsc = (StrComp(Worksheets(j).Name, wkSheet.Name) < 0) Then
                Set wkSheet = Worksheets(j)
            End If
        Next
        
        If i <> wkSheet.Index Then
            Call wkSheet.Move(Before:=Worksheets(i))
        End If
    Next
    
    '// �㏈��
    Call Worksheets(1).Activate
    Call gsResumeAppEvents
    
    Call MsgBox(MSG_FINISHED, vbOKOnly, APP_TITLE)
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
    
'    Application.ScreenUpdating = False
'    Set statGauge = New cStatusGauge
'    statGauge.MaxVal = Selection.SpecialCells(xlCellTypeConstants, xlNumbers + xlTextValues).Count
    Call gsSuppressAppEvents
    
    If Selection.Count > 1 Then
        For Each tCell In Selection.SpecialCells(xlCellTypeConstants, xlNumbers + xlTextValues)
            Call psConvValue_sub(tCell, funcFlag)
            
            '// �L�[����
            If GetAsyncKeyState(27) <> 0 Then
                Application.StatusBar = False
                Exit For
            End If
            
'            Call statGauge.addValue(1)
        Next
    Else
        Call psConvValue_sub(ActiveCell, funcFlag)
    End If
    
'    Set statGauge = Nothing
'    Application.ScreenUpdating = True
    Call gsResumeAppEvents
    Exit Sub
ErrorHandler:
    Call gsResumeAppEvents
    If Err.Number = 1004 Then  '// �͈͑I�����������Ȃ��ꍇ
        Call MsgBox(MSG_INVALID_RANGE, vbOKOnly, APP_TITLE)
    Else
        Call gsShowErrorMsgDlg("mdlCommon.psConvValue", Err)
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
'// ���\�b�h�F   �N���b�v�{�[�h�փR�s�[
'// �����F       �I��͈͂��Œ蒷�ɐ��`���ăN���b�v�{�[�h�Ɋi�[����B
'// �����F       �Ȃ�
'// �߂�l�F     �Ȃ�
'// ////////////////////////////////////////////////////////////////////////////
Public Sub psCopyToClipboard()
On Error GoTo ErrorHandler
    Const MAX_LEN   As Integer = 80
    Dim tRange      As udTargetRange
    Dim colLen()    As Integer
    Dim idxRow      As Long
    Dim idxCol      As Integer
    Dim bffText     As String
    Dim rslt        As String
    Dim bffHead     As String
    Dim idxArry     As Integer
    Dim textLen     As Integer
  
    '// �I��͈͂��P��ł��邱�Ƃ̊m�F
    If Selection.Areas.Count > 1 Then
        Call MsgBox(MSG_TOO_MANY_RANGE, vbOKOnly, APP_TITLE)
        Exit Sub
    End If
  
    tRange = gfGetTargetRange(ActiveSheet, Selection)
  
    If (tRange.minRow > tRange.maxRow) Or (tRange.minCol > tRange.maxCol) Then
        Call MsgBox(MSG_INVALID_RANGE, vbOKOnly, APP_TITLE)
        Exit Sub
    End If
  
    '// �Z���̒������m�F colLen�Ɋi�[
    idxArry = 0
    For idxCol = tRange.minCol To tRange.maxCol
        ReDim Preserve colLen(idxArry + 1)
        For idxRow = tRange.minRow To tRange.maxRow
            textLen = LenB(StrConv(WorksheetFunction.Clean(Cells(idxRow, idxCol).Text), vbFromUnicode))
            If textLen > colLen(idxArry) Then
              colLen(idxArry) = textLen
            End If
        Next
        colLen(idxArry) = IIf(colLen(idxArry) = 0, 1, colLen(idxArry))
        colLen(idxArry) = IIf(colLen(idxArry) > MAX_COL_LEN, MAX_COL_LEN, colLen(idxArry))  '// 80�o�C�g�ȏ�̒����͐؂�̂�
        idxArry = idxArry + 1
    Next
  
    For idxRow = tRange.minRow To tRange.maxRow
        For idxCol = 0 To tRange.Columns - 1
            bffText = Trim(WorksheetFunction.Clean(Cells(idxRow, idxCol + tRange.minCol).Text)) '// ���s�폜��Trim�B�O���̃g�����͐��l�^�̏ꍇ�̕����p�󔒏����̂��ߕK�v
            bffText = StrConv(LeftB(StrConv(bffText, vbFromUnicode), 80), vbUnicode) '// �ő啶�����ȏ�𑫂���
            textLen = LenB(Trim$(StrConv(bffText, vbFromUnicode)))
            If textLen > MAX_LEN Then    '// 80�����ȏ�͐؂�̂�
                bffText = StrConv(LeftB(StrConv(bffText, vbFromUnicode), colLen(idxCol)), vbUnicode)
            ElseIf IsNumeric(bffText) Or IsDate(bffText) Or pfIsPercentage(bffText) Then    '// ���l�A���t�͉E��
                bffText = Space(colLen(idxCol) - LenB(StrConv(bffText, vbFromUnicode))) & bffText
            Else
                bffText = bffText & Space(colLen(idxCol) - LenB(StrConv(bffText, vbFromUnicode)))
            End If
            rslt = rslt & bffText & Space(1)
        Next
        rslt = Left(rslt, Len(rslt) - 1) & vbCrLf
    Next

    '// �擪�Ɩ����Ɍr����ǉ�
    For idxCol = 0 To tRange.Columns - 1
        bffHead = bffHead & String(colLen(idxCol), "-") & IIf(idxCol = tRange.Columns - 1, vbCrLf, " ")
    Next
    rslt = bffHead & rslt & bffHead
    
    '// �N���b�v�{�[�h�փR�s�[ ��Win10����DataObject�����삵�Ȃ��Ȃ邽�߁A�����SetClip�ɒu������
    Call psSetClip(rslt)
'  Set objData = New DataObject
'  Call objData.SetText(rslt)
'  Call objData.PutInClipboard
'  Set objData = Nothing
  
    Exit Sub
ErrorHandler:
    Call gsShowErrorMsgDlg("mdlCommon.psCopyToClipboard", Err)
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// ���\�b�h�F   �N���b�v�{�[�h��Markdown�`���ŃR�s�[
'// �����F       �I��͈͂��}�[�N�_�E���`���ŃN���b�v�{�[�h�Ɋi�[����B
'// �����F       �Ȃ�
'// �߂�l�F     �Ȃ�
'// ////////////////////////////////////////////////////////////////////////////
Public Sub psCopyToCB_Markdown()
On Error GoTo ErrorHandler
    Dim tRange      As udTargetRange
    Dim colLen()    As Integer
    Dim idxRow      As Long
    Dim idxCol      As Integer
    Dim bffText     As String
    Dim rslt        As String
    Dim bffHead     As String
    Dim idxArry     As Integer
    Dim textLen     As Integer
  
    '// �I��͈͂��P��ł��邱�Ƃ̊m�F
    If Selection.Areas.Count > 1 Then
        Call MsgBox(MSG_TOO_MANY_RANGE, vbOKOnly, APP_TITLE)
        Exit Sub
    End If
  
    tRange = gfGetTargetRange(ActiveSheet, Selection)
  
    If (tRange.minRow > tRange.maxRow) Or (tRange.minCol > tRange.maxCol) Then
        Call MsgBox(MSG_INVALID_RANGE, vbOKOnly, APP_TITLE)
        Exit Sub
    End If
    
    '// �w�b�_�̏o��
    rslt = "|"
    For idxCol = 0 To tRange.Columns - 1
        rslt = rslt & " " & Replace(Cells(tRange.minRow, idxCol + tRange.minCol).Text, vbLf, "<br>") & " |"
    Next
    rslt = rslt & vbCrLf & "|"
    For idxCol = 0 To tRange.Columns - 1
        Select Case Cells(tRange.minRow, idxCol + tRange.minCol).HorizontalAlignment
            Case xlRight
                rslt = rslt & " " & "-: |"
            Case xlCenter
                rslt = rslt & " " & ":-: |"
            Case Else
                rslt = rslt & " " & "- |"
        End Select
    Next
    
    '// �f�[�^�s�̏o��
    For idxRow = tRange.minRow + 1 To tRange.maxRow
        rslt = rslt & vbCrLf & "|"
        For idxCol = 0 To tRange.Columns - 1
            rslt = rslt & " " & Replace(Cells(idxRow, idxCol + tRange.minCol).Text, vbLf, "<br>") & " |"
        Next
    Next
    
    '// �N���b�v�{�[�h�փR�s�[ ��Win10����DataObject�����삵�Ȃ��Ȃ邽�߁A�����SetClip�ɒu������
    Call psSetClip(rslt)
    
    Exit Sub
ErrorHandler:
    Call gsShowErrorMsgDlg("mdlCommon.psCopyToClipboard_MarkDown", Err)
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// �����F       Win10 ���� DataObject.PutInClipboard �������Ȃ��Ȃ������߁A�����Ƃ��ăe�L�X�g�{�b�N�X���o�R���ăR�s�[
'// �����F       �R�s�[�Ώە�����
'// ////////////////////////////////////////////////////////////////////////////
Sub psSetClip(bffText As String)
    With CreateObject("Forms.TextBox.1")
        .MultiLine = True
        .Text = bffText
        .SelStart = 0
        .SelLength = .TextLength
        .Copy
    End With
    DoEvents    '// ��Q����̂��߁A��xOS�ɏ�����߂��i�Č����Ⴂ���߂��̑Ώ����ǂ����͊m�ؖ����j
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// �����F       �����̕����񂪃p�[�Z���g�`�����𔻒肷��
'// �����F       �R�s�[�Ώە�����
Function pfIsPercentage(bffText As String) As Boolean
    If bffText = BLANK Then
        pfIsPercentage = False
    ElseIf Right(bffText, 1) = "%" And IsNumeric(Left(bffText, Len(bffText) - 1)) Then
        pfIsPercentage = True
    Else
        pfIsPercentage = False
    End If
End Function


'// ////////////////////////////////////////////////////////////////////////////
'// ���\�b�h�F   �I��͈͐ݒ�i�F�ɂ��j
'// �����F       �A�N�e�B�u�Z���Ɠ����F�̃Z����I��͈͂ɐݒ肷��
'// ////////////////////////////////////////////////////////////////////////////
'Private Sub psSetupSelection_color(colorMode As String)
'    Dim targetCell  As Range
'    Dim rgbColor    As Long
'    Dim rslt        As Range
'
'    '// �����ݒ�
'    If colorMode = "B" Then
'        rgbColor = ActiveCell.Interior.Color
'    Else
'        rgbColor = ActiveCell.Font.Color
'    End If
'
'    '// �f�t�H���g�F�̏ꍇ�̓L�����Z���𑣂�
'    If (colorMode = "B" And rgbColor = 16777215) _
'      Or (colorMode = "F" And rgbColor = 0) Then
'        If MsgBox(MSG_SEL_DEFAULT_COLOR, vbOKCancel, APP_TITLE) = vbCancel Then
'            Exit Sub
'        End If
'    End If
'
'    Application.ScreenUpdating = False
'    Set rslt = ActiveCell
'    For Each targetCell In ActiveSheet.UsedRange
'        If colorMode = "B" Then
'            If targetCell.Interior.Color = rgbColor Then  '// �Z���w�i�F�̔���
'                Set rslt = Union(rslt, targetCell)
'            End If
'        Else
'            If targetCell.Font.Color = rgbColor Then      '// �t�H���g�F�̔���
'                Set rslt = Union(rslt, targetCell)
'            End If
'        End If
'    Next
'
'    Call rslt.Select
'    Application.ScreenUpdating = True
'End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// ���\�b�h�F   �r���`��i�w�b�_�j
'// �����F       �w�b�_���̌r����`�悷��i���j
'// ////////////////////////////////////////////////////////////////////////////
Public Sub gsDrawLine_Header()
    Dim baseRow As Long     '// �I��̈�̊J�n�ʒu
    Dim baseCol As Integer  '// �I��̈�̊J�n�ʒu
    Dim selRows As Long     '// �I��̈�̍s��
    Dim selCols As Integer  '// �I��̈�̗�
    Dim idxRow  As Long
    Dim idxCol  As Integer
    Dim offRow  As Long
    Dim offCol  As Integer
    Dim childRange As Range
    
    '// ���O�`�F�b�N�i�A�N�e�B�u�V�[�g�ی�A�I���^�C�v���Z���j
    If Not gfPreCheck(protectCont:=True, selType:=TYPE_RANGE) Then
        Exit Sub
    End If
'    Application.ScreenUpdating = False
    Call gsSuppressAppEvents
    
    For Each childRange In Selection.Areas
        '// �r�����N���A
        childRange.Borders.LineStyle = xlNone
        childRange.Borders(xlDiagonalDown).LineStyle = xlNone
        childRange.Borders(xlDiagonalUp).LineStyle = xlNone
        
        '// �I��͈͂̊J�n�E�I���ʒu�擾
        baseRow = childRange.Row
        baseCol = childRange.Column
        selRows = childRange.Rows.Count
        selCols = childRange.Columns.Count
        
        For idxRow = baseRow To baseRow + selRows
            For idxCol = baseCol To baseCol + selCols
                offRow = 0
                offCol = 0
                If (Cells(idxRow, idxCol).Text <> BLANK) Or ((idxRow = baseRow) And (idxCol = baseCol)) Then
                    For offRow = idxRow To baseRow + selRows - 1
                        If (offRow = idxRow) Or Cells(offRow, idxCol).Value = BLANK Then
                            Cells(offRow, idxCol).Borders(xlEdgeLeft).LineStyle = xlContinuous
                        Else
                            Exit For
                        End If
                    Next
                    For offCol = idxCol To baseCol + selCols - 1
                        If (offCol = idxCol) Or Cells(idxRow, offCol).Text = BLANK Then
                            Cells(idxRow, offCol).Borders(xlEdgeTop).LineStyle = xlContinuous
                            If Cells(idxRow, offCol).Borders(xlEdgeRight).LineStyle = xlContinuous Then
                                Exit For
                            End If
                        Else
                            Exit For
                        End If
                    Next
                End If
                '// �ő��ɒB�����ꍇ�͏I��
                If idxCol = Columns.Count Then
                    Exit For
                End If
            Next
            '// �ő�s�ɒB�����ꍇ�͏I��
            If idxRow = Rows.Count Then
                Exit For
            End If
        Next
        
        With childRange.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With
        With childRange.Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With
        'childRange.Interior.ColorIndex = COLOR_HEADER
    Next
'    Application.ScreenUpdating = True
    Call gsResumeAppEvents
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// ���\�b�h�F   �r���`��i�w�b�_�j�F�c
'// �����F       �w�b�_���̌r����`�悷��i�c�j
'// �����F       �Ȃ�
'// �߂�l�F     �Ȃ�
'// ////////////////////////////////////////////////////////////////////////////
Public Sub gsDrawLine_Header_Vert()
    Dim baseRow As Long     '// �I��̈�̊J�n�ʒu
    Dim baseCol As Integer  '// �I��̈�̊J�n�ʒu
    Dim selRows As Long     '// �I��̈�̍s��
    Dim selCols As Integer  '// �I��̈�̗�
    Dim idxRow  As Long
    Dim idxCol  As Integer
    Dim offRow  As Long
    Dim offCol  As Integer
    Dim childRange As Range
  
    '// ���O�`�F�b�N�i�A�N�e�B�u�V�[�g�ی�A�I���^�C�v���Z���j
    If Not gfPreCheck(protectCont:=True, selType:=TYPE_RANGE) Then
        Exit Sub
    End If
'    Application.ScreenUpdating = False
    Call gsSuppressAppEvents
    
    For Each childRange In Selection.Areas
        '// �r�����N���A
        childRange.Borders.LineStyle = xlNone
        childRange.Borders(xlDiagonalDown).LineStyle = xlNone
        childRange.Borders(xlDiagonalUp).LineStyle = xlNone
        
        '// �I��͈͂̊J�n�E�I���ʒu�擾
        baseRow = childRange.Row
        baseCol = childRange.Column
        selRows = childRange.Rows.Count
        selCols = childRange.Columns.Count
      
        For idxCol = baseCol To baseCol + selCols
            For idxRow = baseRow To baseRow + selRows
                offRow = 0
                offCol = 0
                If (Cells(idxRow, idxCol).Value <> BLANK) Or ((idxRow = baseRow) And (idxCol = baseCol)) Then
                    For offCol = idxCol To baseCol + selCols - 1
                        If (offCol = idxCol) Or Cells(idxRow, offCol).Value = BLANK Then
                            Cells(idxRow, offCol).Borders(xlEdgeTop).LineStyle = xlContinuous
                        Else
                            Exit For
                        End If
                    Next
                    For offRow = idxRow To baseRow + selRows - 1
                        If (offRow = idxRow) Or Cells(offRow, idxCol).Value = BLANK Then
                            Cells(offRow, idxCol).Borders(xlEdgeLeft).LineStyle = xlContinuous
                            If Cells(offRow, idxCol).Borders(xlEdgeBottom).LineStyle = xlContinuous Then
                                Exit For
                            End If
                        Else
                            Exit For
                        End If
                    Next
                End If
                '// �ő�s�ɒB�����ꍇ�͏I��
                If idxRow = Rows.Count Then
                    Exit For
                End If
            Next
            '// �ő��ɒB�����ꍇ�͏I��
            If idxCol = Columns.Count Then
                Exit For
            End If
        Next
    
        With childRange.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With
        With childRange.Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With
    
    'childRange.Interior.ColorIndex = COLOR_HEADER
    Next
    
    Call gsResumeAppEvents
'    Application.ScreenUpdating = True
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// ���\�b�h�F   �r���`��i�f�[�^�j
'// �����F       �f�[�^���̌r����`�悷��
'//              �I��͈͎��ӕ���xlThin�A������xlHairline�ŕ`�悷��
'// �����F       �Ȃ�
'// �߂�l�F     �Ȃ�
'// ////////////////////////////////////////////////////////////////////////////
Public Sub gsDrawLine_Data()
    '// ���O�`�F�b�N�i�A�N�e�B�u�V�[�g�ی�A�I���^�C�v���Z���j
    If Not gfPreCheck(protectCont:=True, selType:=TYPE_RANGE) Then
        Exit Sub
    End If
    
    Call gsSuppressAppEvents
'    Application.ScreenUpdating = False
    
    '// V2.0���A�����ʒu�͌���̂܂܂Ƃ���iSelection.VerticalAlignment�͕ύX���Ȃ��j�悤�ύX�B
    '// �����ʒu���㕔�ɐݒ�
'    Selection.VerticalAlignment = xlTop
    
    '// �r���`��
    Selection.Borders.LineStyle = xlContinuous
    Selection.Borders.Weight = xlThin
    
    If Selection.Columns.Count > 1 Then
        With Selection.Borders(xlInsideVertical)
            .LineStyle = xlContinuous
            .Weight = xlHairline
        End With
    End If
    
    If Selection.Rows.Count > 1 Then
        With Selection.Borders(xlInsideHorizontal)
            .LineStyle = xlContinuous
            .Weight = xlHairline
        End With
    End If
    
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone

'    Application.ScreenUpdating = True
    Call gsResumeAppEvents
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// ���\�b�h�F    �y�[�W�ݒ�(�w�b�_�E�t�b�^)
'// �����F        �y�[�W�ݒ���s��
'// �����F        wksheet: ���[�N�V�[�g
'// ////////////////////////////////////////////////////////////////////////////
Public Sub gsPageSetup_Header(wkSheet As Worksheet)
'// $mod �v�����^���Ȃ��ꍇ�̖����I�ȃG���[�́H
On Error Resume Next
    '// �v�����^�̐ݒ�
    With wkSheet.PageSetup
        '// �w�b�_  ���쐬�҂�\������ꍇ�͉E�w�b�_�̃R�����g�A�E�g�����g�p�B
        .LeftHeader = HED_LEFT
        .CenterHeader = HED_CENTER
        .RightHeader = HED_RIGHT
        '// .RightHeader = "&""" & APP_FONT & ",�W��""&8�쐬��:" & Application.UserName & IIf(Application.OrganizationName = BLANK, BLANK, "@" & Application.OrganizationName)
        '// �t�b�^
        .LeftFooter = FOT_LEFT
        .CenterFooter = FOT_CENTER
        .RightFooter = FOT_RIGHT
    End With
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// ���\�b�h�F    �y�[�W�ݒ�(�}�[�W��)
'// �����F        �}�[�W���̐ݒ���s��
'// �����F        wksheet: ���[�N�V�[�g
'// ////////////////////////////////////////////////////////////////////////////
Public Sub gsPageSetup_Margin(wkSheet As Worksheet)
On Error Resume Next
    '// �v�����^�̐ݒ�
    With wkSheet.PageSetup
        '// �}�[�W��
        .LeftMargin = Application.InchesToPoints(MRG_LEFT)
        .RightMargin = Application.InchesToPoints(MRG_RIGHT)
        .TopMargin = Application.InchesToPoints(MRG_TOP)
        .BottomMargin = Application.InchesToPoints(MRG_BOTTOM)
        .HeaderMargin = Application.InchesToPoints(MRG_HEADER)
        .FooterMargin = Application.InchesToPoints(MRG_FOOTER)
    End With
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// ���\�b�h�F    �y�[�W�ݒ�(�r��)
'// �����F        �r����`�悷��
'// �����F        wksheet: ���[�N�V�[�g
'//               headerLines: �w�b�_�s��
'// ////////////////////////////////////////////////////////////////////////////
Public Sub gsPageSetup_Lines(wkSheet As Worksheet, headerLines As Integer)
    '// �r����`��
    Call wkSheet.UsedRange.Select
    Call mdlCommon.gsDrawLine_Data
  
    '// �w�b�_�̏C��
    If headerLines > 0 Then
        Call wkSheet.Range(wkSheet.Cells(1, 1), wkSheet.Cells(headerLines, wkSheet.UsedRange.Columns.Count)).Select
        Call mdlCommon.gsDrawLine_Header
    
        '// �w�b�_�����ŃE�B���h�E�g���Œ�
        Call wkSheet.Cells(headerLines + 1, 1).Select
        ActiveWindow.FreezePanes = True
    End If
    
    Call wkSheet.Cells(1, 1).Select
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// ���\�b�h�F   �L���͈͐ݒ�
'// �����F       �I��͈͂ƒl�̐ݒ肳��Ă���͈͂��r���A�L���͈͂��擾����
'// �����F       wksheet: ���[�N�V�[�g
'//              selRange: �I��͈�
'// �߂�l�F     �␳��̑I��͈�
'// ////////////////////////////////////////////////////////////////////////////
Public Function gfGetTargetRange(wkSheet As Worksheet, selRange As Range) As udTargetRange
    Dim rslt  As udTargetRange
    
    rslt.minRow = selRange.Row
    rslt.minCol = selRange.Column
    rslt.maxRow = IIf(wkSheet.UsedRange.Row + wkSheet.UsedRange.Rows.Count < selRange.Row + selRange.Rows.Count, wkSheet.UsedRange.Row + wkSheet.UsedRange.Rows.Count - 1, selRange.Row + selRange.Rows.Count - 1)
    rslt.maxCol = IIf(wkSheet.UsedRange.Column + wkSheet.UsedRange.Columns.Count < selRange.Column + selRange.Columns.Count, wkSheet.UsedRange.Column + wkSheet.UsedRange.Columns.Count - 1, selRange.Column + selRange.Columns.Count - 1)
    rslt.Rows = rslt.maxRow - rslt.minRow + 1
    rslt.Columns = rslt.maxCol - rslt.minCol + 1
    
    gfGetTargetRange = rslt
End Function


'// ////////////////////////////////////////////////////////////////////////////
'// ���\�b�h�F   �n�C�p�[�����N�̐ݒ�
'// �����F       �I��͈͂̃n�C�p�[�����N��ݒ肷��
'//              �W���@�\�̃n�C�p�[�����N�ݒ�ł̓e�L�X�g�������ς�邽�߁A�ݒ�O�̏�����ێ�����
'// �����F       �Ȃ�
'// �߂�l�F     �Ȃ�
'// ////////////////////////////////////////////////////////////////////////////
Private Sub psSetHyperLink()
    Dim tRange    As udTargetRange
    Dim childRange As Range
    Dim idxRow    As Long
    Dim idxCol    As Integer
    Dim fontName  As String
    Dim fontSize  As String
    Dim fontBold  As Boolean
    Dim fontItlic As Boolean
    Dim fontColor As Double
    
    '// ���O�`�F�b�N�i�A�N�e�B�u�V�[�g�ی�A�I���^�C�v���Z���j
    If Not gfPreCheck(protectCont:=True, selType:=TYPE_RANGE) Then
        Exit Sub
    End If
    
    Call gsSuppressAppEvents
'    Application.ScreenUpdating = False
    
    For Each childRange In Selection.Areas
        tRange = gfGetTargetRange(ActiveSheet, childRange)
        For idxRow = tRange.minRow To tRange.maxRow
            For idxCol = tRange.minCol To tRange.maxCol
                If Trim(Cells(idxRow, idxCol).Text) <> BLANK Then
                    fontName = Cells(idxRow, idxCol).Font.Name
                    fontSize = Cells(idxRow, idxCol).Font.Size
                    fontBold = Cells(idxRow, idxCol).Font.Bold
                    fontItlic = Cells(idxRow, idxCol).Font.Italic
                    fontColor = Cells(idxRow, idxCol).Font.Color
                    Call Cells(idxRow, idxCol).Hyperlinks.Add(Anchor:=Cells(idxRow, idxCol), Address:=Cells(idxRow, idxCol).Text)
                    Cells(idxRow, idxCol).Font.Name = fontName
                    Cells(idxRow, idxCol).Font.Size = fontSize
                    Cells(idxRow, idxCol).Font.Bold = fontBold
                    Cells(idxRow, idxCol).Font.Italic = fontItlic
                    Cells(idxRow, idxCol).Font.Color = fontColor
                End If
            Next
        Next
    
        '// �L�[����
        If GetAsyncKeyState(27) <> 0 Then
            Exit For
        End If
    Next
'    Application.ScreenUpdating = True
    Call gsResumeAppEvents
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// ���\�b�h�F   �n�C�p�[�����N�̍폜
'// �����F       �I��͈͂̃n�C�p�[�����N���폜����
'// �����F       �Ȃ�
'// �߂�l�F     �Ȃ�
'// Excel 2010���_�ŁuHyperLink�̃N���A�v���W����������Ă��邪�A�c�[���Ƃ���UI���c�����ƂƂ����B
'// ////////////////////////////////////////////////////////////////////////////
Private Sub psRemoveHyperLink()
    Dim tRange    As udTargetRange
    Dim idxRow    As Long
    Dim idxCol    As Integer
    Dim fontName  As String
    Dim fontSize  As String
    Dim borderLines(8, 3) As Long
    Dim childRange As Range
    
    '// ���O�`�F�b�N�i�A�N�e�B�u�V�[�g�ی�A�I���^�C�v���Z���j
    If Not gfPreCheck(protectCont:=True, selType:=TYPE_RANGE) Then
        Exit Sub
    End If
    
    Call Selection.ClearHyperlinks
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// ���\�b�h�F   �񕶎���擾
'// �����F       ��̔ԍ��𕶎��\�L�ɕϊ�����
'// �����F       targetVal: ��ԍ�
'// �߂�l�F     ��̕�����\�L
'// ////////////////////////////////////////////////////////////////////////////
Public Function gfGetColIndexString(ByVal targetVal As Integer) As String
    Const ALPHABETS   As Integer = 26
    Dim remainder     As Integer
    Dim rslt          As String
    
    Do
        remainder = IIf((targetVal Mod ALPHABETS) = 0, ALPHABETS, targetVal Mod ALPHABETS)
        rslt = Chr(64 + remainder) & rslt
        targetVal = Int((targetVal - 1) / ALPHABETS)
    Loop Until targetVal < 1
    
    gfGetColIndexString = rslt
End Function


'// ////////////////////////////////////////////////////////////////////////////
'// ���\�b�h�F   �O���[�v�ݒ�i�s�j
'// �����F       �O���[�v�������ݒ肷��B
'// �����F       �Ȃ�
'// �߂�l�F     �Ȃ�
'// ////////////////////////////////////////////////////////////////////////////
Private Sub psSetGroup_Row()
    Dim idxStart    As Long
    Dim idxEnd      As Long
    Dim idxRow      As Long
    Dim idxCol      As Integer
    Dim tRange      As udTargetRange
    Dim childRange  As Range
    
    '// ���O�`�F�b�N�i�A�N�e�B�u�V�[�g�ی�A�I���^�C�v���Z���j
    If Not gfPreCheck(protectCont:=True, selType:=TYPE_RANGE) Then
        Exit Sub
    End If
    
'    Application.ScreenUpdating = False
    Call gsSuppressAppEvents
    
  
    '// �A�E�g���C���̏W�v�ʒu��ύX
    With ActiveSheet.Outline
        .SummaryRow = xlAbove
    End With
  
        '// �O���[�v�ݒ�
        For Each childRange In Selection.Areas
            tRange = gfGetTargetRange(ActiveSheet, childRange)
            
            idxStart = 0
            idxEnd = 0
            idxCol = tRange.minCol
            
            For idxRow = tRange.minRow To tRange.maxRow
                If idxStart = 0 Then
                    idxStart = idxRow + 1
                    idxEnd = idxRow + 1
                ElseIf Trim(Cells(idxRow, idxCol).Text) = BLANK Then
                    idxEnd = idxRow
                ElseIf Trim(Cells(idxRow - 1, idxCol).Text) = BLANK Then
                    Range(Cells(idxStart, 1), Cells(idxEnd, 1)).Rows.Group
                    idxStart = idxRow + 1
                    idxEnd = idxRow + 1
                Else
                    idxStart = idxRow + 1
                    idxEnd = idxRow + 1
                End If
            Next
            If idxStart < idxEnd Then
                Range(Cells(idxStart, 1), Cells(idxEnd, 1)).Rows.Group
            End If
      Next
      
      Call gsResumeAppEvents
'      Application.ScreenUpdating = True
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// ���\�b�h�F   �O���[�v�ݒ�i��j
'// �����F       �O���[�v�������ݒ肷��B
'// �����F       �Ȃ�
'// �߂�l�F     �Ȃ�
'// ////////////////////////////////////////////////////////////////////////////
Private Sub psSetGroup_Col()
    Dim idxStart    As Long
    Dim idxEnd      As Long
    Dim idxRow      As Long
    Dim idxCol      As Integer
    Dim tRange      As udTargetRange
    Dim childRange  As Range
    
    '// ���O�`�F�b�N�i�A�N�e�B�u�V�[�g�ی�A�I���^�C�v���Z���j
    If Not gfPreCheck(protectCont:=True, selType:=TYPE_RANGE) Then
        Exit Sub
    End If
    
    Call gsSuppressAppEvents
'    Application.ScreenUpdating = False
    
    '// �A�E�g���C���̏W�v�ʒu��ύX
    With ActiveSheet.Outline
        .SummaryColumn = xlLeft
    End With
    
    '// �O���[�v�ݒ�
    For Each childRange In Selection.Areas
        tRange = gfGetTargetRange(ActiveSheet, childRange)
        
        idxStart = 0
        idxEnd = 0
        idxRow = tRange.minRow
        
        For idxCol = tRange.minCol To tRange.maxCol
            If idxStart = 0 Then
                idxStart = idxCol + 1
                idxEnd = idxCol + 1
            ElseIf Trim(Cells(idxRow, idxCol).Text) = BLANK Then
                idxEnd = idxCol
            ElseIf Trim(Cells(idxRow, idxCol - 1).Text) = BLANK Then
                Range(Cells(1, idxStart), Cells(1, idxEnd)).Columns.Group
                idxStart = idxCol + 1
                idxEnd = idxCol + 1
            Else
                idxStart = idxCol + 1
                idxEnd = idxCol + 1
            End If
        Next
        If idxStart < idxEnd Then
            Range(Cells(1, idxStart), Cells(1, idxEnd)).Columns.Group
        End If
    Next
    
'    Application.ScreenUpdating = True
    Call gsResumeAppEvents
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// ���\�b�h�F   �l�̏d����r�����Ĉꗗ�i�J�E���g�j
'// �����F       �d���l��r������B
'// ////////////////////////////////////////////////////////////////////////////
Private Sub psDistinctVals()
    Dim idxRow      As Long
    Dim idxCol      As Integer
    Dim tRange      As udTargetRange
    
    Dim bff         As Variant
    Dim dict        As Object
    Dim keyString   As String
    Dim keyArray()  As String
    Dim resultSheet As Worksheet
    
    '// �Z�����I������Ă��邱�Ƃ��`�F�b�N
    If TypeName(Selection) <> TYPE_RANGE Then
        Call MsgBox(MSG_NOT_RANGE_SELECT, vbOKOnly, APP_TITLE)
        Exit Sub
    End If
    
    '// �`�F�b�N
    If Selection.Areas.Count > 1 Then
        Call MsgBox(MSG_TOO_MANY_RANGE, vbOKOnly, APP_TITLE)
        Exit Sub
    End If
    
    Call gsSuppressAppEvents
'    Application.ScreenUpdating = False
    
    tRange = gfGetTargetRange(ActiveSheet, Selection)
    
    bff = Selection.Areas(1)
    Set dict = CreateObject("Scripting.Dictionary")
    
    For idxRow = 1 To tRange.maxRow - tRange.minRow + 1
        '// �s�̃Z�����������ĕ���������
        keyString = BLANK
        For idxCol = 1 To tRange.maxCol - tRange.minCol + 1
            If Not IsError(bff(idxRow, idxCol)) Then
                keyString = keyString & Chr(127) & bff(idxRow, idxCol)
            End If
        Next
        
        If Not dict.Exists(keyString) Then
            Call dict.Add(keyString, "1")
        Else
            dict.Item(keyString) = CStr(CLng(dict.Item(keyString)) + 1)
        End If
    Next
    
    '// ���ʏo��
    Call Workbooks.Add
    Set resultSheet = ActiveWorkbook.ActiveSheet
    
    '// �w�b�_�̐ݒ�B�u�J�E���g�v�̃w�b�_�ʒu�����킹�邽�߁AHDR_DISTINCT����"@"��񐔂ɍ��킹��Replace����
    Call gsDrawResultHeader(resultSheet, Replace(HDR_DISTINCT, "@", String(tRange.Columns, ";")), 1)
    
    '// �L�[�̔z���variant�Ɋi�[
    bff = dict.Keys
    
    For idxRow = 0 To dict.Count - 1
        keyArray = Split(bff(idxRow), Chr(127))  '// split�͓Y�����P����J�n�̎d�l�H
        For idxCol = 1 To UBound(keyArray)
            resultSheet.Cells(idxRow + 2, idxCol).Value = keyArray(idxCol)
        Next
        
        resultSheet.Cells(idxRow + 2, tRange.maxCol - tRange.minCol + 2).Value = dict.Item(bff(idxRow))
    Next
    
    '//�t�H���g
    Call resultSheet.Cells.Select
    Selection.Font.Name = APP_FONT
    Selection.Font.Size = APP_FONT_SIZE
    
    '// �r���`��
    Call gsPageSetup_Lines(resultSheet, 1)
    
    '// ����Ƃ��ɕۑ������߂Ȃ�
    ActiveWorkbook.Saved = True
    
    Set dict = Nothing
    
'    Application.ScreenUpdating = True
    Call gsResumeAppEvents

End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// ���\�b�h�F   �l���K�w���ɕ␳����
'// �����F       �d���l���K�w���ɕ␳����
'// ////////////////////////////////////////////////////////////////////////////
Private Sub psGroupVals()
    Dim idxRow        As Long
    Dim idxCol        As Integer
    Dim tRange        As udTargetRange
    Dim aryIdx        As Integer
    Dim aryLastVal(8) As String
    
    '// �Z�����I������Ă��邱�Ƃ��`�F�b�N
    If TypeName(Selection) <> TYPE_RANGE Then
        Call MsgBox(MSG_NOT_RANGE_SELECT, vbOKOnly, APP_TITLE)
        Exit Sub
    End If
    
    '// �`�F�b�N
    If Selection.Areas.Count > 1 Then
        Call MsgBox(MSG_TOO_MANY_RANGE, vbOKOnly, APP_TITLE)
        Exit Sub
    ElseIf Selection.Columns.Count > 8 Then
        Call MsgBox(MSG_TOO_MANY_COLS_8, vbOKOnly, APP_TITLE)
        Exit Sub
    End If
    
    Call gsSuppressAppEvents
'    Application.ScreenUpdating = False
    
    tRange = gfGetTargetRange(ActiveSheet, Selection)
    
    For idxRow = tRange.minRow To tRange.maxRow
        For idxCol = tRange.minCol To tRange.maxCol
            If (aryLastVal(idxCol - tRange.minCol) = BLANK) Or (aryLastVal(idxCol - tRange.minCol) <> Cells(idxRow, idxCol).Text) Then
                '// ���O�̒l���قȂ�ꍇ (�� ���O�̒l���u�����N�̏ꍇ)
                '// �z���̃��x���̒��O�̒l���N���A
                For aryIdx = tRange.Columns To idxCol Step -1
                    aryLastVal(aryIdx - 1) = BLANK
                Next
                aryLastVal(idxCol - tRange.minCol) = Cells(idxRow, idxCol).Text
            Else
                Cells(idxRow, idxCol).Value = BLANK
            End If
        Next
    Next
    
'    Application.ScreenUpdating = True
    Call gsResumeAppEvents
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// ���\�b�h�F   VLookup�̃}�X�^�̈�Ƃ��ăR�s�[
'// �����F       �I��̈��\�������������ϐ��Ɋi�[����
'// ////////////////////////////////////////////////////////////////////////////
Private Sub psVLookupCopy()
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
    Call gsShowErrorMsgDlg("mdlCommon.psVLookupPaste", Err)
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// ���\�b�h�F   �Z��������擾
'// �����F       text �܂��� value �v���p�e�B�̒l��Ԃ�
'//              ������(@)�̏ꍇ�ɂ� .Text ��߂��A����ȊO�̏ꍇ�� $todo
'// �����F       targetCell: �ΏۃZ��
'// �߂�l�F     �����̃Z���̒l�itext/value�v���p�e�B�j
'// ////////////////////////////////////////////////////////////////////////////
Public Function gfGetTextVal(targetCell As Range) As String
    gfGetTextVal = IIf(targetCell.NumberFormat = "@", targetCell.Value, targetCell.Text)
End Function


'// ////////////////////////////////////////////////////////////////////////////
'// ���\�b�h�F   �R���{�{�b�N�X�ݒ�
'// �����F       ������CSV���������ɁA�R���{�{�b�N�X�̒l��ݒ肷��B
'// �����F       targetCombo: �ΏۃR���{�{�b�N�X
'//              propertyStr: �ݒ�l�i{�L�[},{�\��������};{�L�[},{�\��������}...�j
'//              defaultIdx: �����l
'// ////////////////////////////////////////////////////////////////////////////
Public Sub gsSetCombo(targetCombo As ComboBox, propertyStr As String, defaultIdx As Integer)
    Dim lineStr()     As String   '// �ݒ�l�̕����񂩂�A�e�s���i�[�i;��؂�j
    Dim colStr()      As String   '// �e�s�̕����񂩂�A�񂲂Ƃ̒l���i�[�i,��؂�j
    Dim idxCnt        As Integer
    
    lineStr = Split(propertyStr, ";")     '//�ݒ�l�̕�������A�s���ɕ���
    
    Call targetCombo.Clear
    For idxCnt = 0 To UBound(lineStr)
        colStr = Split(lineStr(idxCnt), ",")   '//�s�̕�������A�J�������̕�����ɕ���
        Call targetCombo.AddItem(Trim(colStr(0)))
        targetCombo.List(idxCnt, 1) = Trim(colStr(1))
    Next
    
    targetCombo.ListIndex = defaultIdx    '// �����l��ݒ�
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// ���\�b�h�F   �t�H���_�I���_�C�A���O�\��
'// �����F       �t�H���_�I���_�C�A���O��\������B
'// �����F       lngHwnd �E�B���h�E�n���h��
'//              strReturnPath �w�肳�ꂽ�t�H���_�̃p�X������
'// �߂�l�F     True:����  False:���s(�L�����Z����I�������ꍇ�܂�)
'// ////////////////////////////////////////////////////////////////////////////
Public Function gfShowSelectFolder(ByVal lngHwnd As Long, ByRef strReturnPath) As Boolean
    Dim lngRet        As Long
    Dim lngReturnCode As Long
    Dim strPath       As String
    Dim biInfo        As BROWSEINFO
    
    lngRet = False
    
    '//������̈�̊m��
    strPath = String(MAX_PATH + 1, Chr(0))
    
    ' �\���̂̏�����
    biInfo.hwndOwner = lngHwnd
    biInfo.lpszTitle = APP_TITLE
    biInfo.ulFlags = BIF_RETURNONLYFSDIRS
    
    '// �t�H���_�I���_�C�A���O�̕\��
    lngReturnCode = apiSHBrowseForFolder(biInfo)
    
    If lngReturnCode <> 0 Then
        Call apiSHGetPathFromIDList(lngReturnCode, strPath)
        strReturnPath = Left(strPath, InStr(strPath, vbNullChar) - 1)
        gfShowSelectFolder = True
    Else
        gfShowSelectFolder = False
    End If
End Function


'// ////////////////////////////////////////////////////////////////////////////
'// ���\�b�h�F   �������s�O�`�F�b�N�i�ėp�j
'// �����F       �e�����̎��s�O�`�F�b�N���s��
'// �����F
'// �߂�l�F     True:����  False:���s
'// ////////////////////////////////////////////////////////////////////////////
Public Function gfPreCheck(Optional protectCont As Boolean = False, _
                            Optional protectBook As Boolean = False, _
                            Optional selType As String = BLANK, _
                            Optional selAreas As Integer = 0, _
                            Optional selCols As Integer = 0) As Boolean
  
    gfPreCheck = True
    
    If ActiveSheet Is Nothing Then                              '// �V�[�g�i�u�b�N�j���J����Ă��邩
        Call MsgBox(MSG_NO_BOOK, vbOKOnly, APP_TITLE)
        gfPreCheck = False
        Exit Function
    End If
    
    If protectCont And ActiveSheet.ProtectContents Then         '// �A�N�e�B�u�V�[�g���ی삳��Ă��邩
        Call MsgBox(MSG_SHEET_PROTECTED, vbOKOnly, APP_TITLE)
        gfPreCheck = False
        Exit Function
    End If
    
    If protectBook And ActiveWorkbook.ProtectStructure Then     '// �u�b�N���ی삳��Ă��邩
        Call MsgBox(MSG_BOOK_PROTECTED, vbOKOnly, APP_TITLE)
        gfPreCheck = False
        Exit Function
    End If
    
    '// �I��͈͂̃^�C�v���`�F�b�N
    Select Case selType
        Case TYPE_RANGE
            If TypeName(Selection) <> TYPE_RANGE Then
                Call MsgBox(MSG_NOT_RANGE_SELECT, vbOKOnly, APP_TITLE)
                gfPreCheck = False
                Exit Function
            End If
        Case TYPE_SHAPE
            If Not VarType(ActiveWindow.Selection) = vbObject Then
                Call MsgBox(MSG_SHAPE_NOT_SELECTED, vbOKOnly, APP_TITLE)
                gfPreCheck = False
                Exit Function
            End If
        Case BLANK
            '// null
    End Select
    
    '// �I��͈̓J�E���g
    If selAreas > 1 Then
        If Selection.Areas.Count > selAreas Then
            Call MsgBox(MSG_TOO_MANY_RANGE, vbOKOnly, APP_TITLE)
            gfPreCheck = False
            Exit Function
        End If
    End If
    
    '// �I��͈̓Z���J�E���g
    If selCols > 1 Then
        If Selection.Columns.Count > selCols Then
            Call MsgBox(MSG_TOO_MANY_COLS_8, vbOKOnly, APP_TITLE)
            gfPreCheck = False
            Exit Function
        End If
    End If
End Function


'// ////////////////////////////////////////////////////////////////////////////
'// ���\�b�h�F   ���ʃV�[�g �w�b�_�`��
'// �����F       �����̃w�b�_��������V�[�g�ɏo�͂���
'// �����F       wkSheet �ΏۃV�[�g
'//              headerStr  �o�͂��镶����
'//              idxRow  �o�͂���s
'// ////////////////////////////////////////////////////////////////////////////
Public Sub gsDrawResultHeader(wkSheet As Worksheet, headerStr As String, idxRow As Integer)
    Dim idxCol      As Integer
    Dim aryString() As String
    
    aryString = Split(headerStr, ";")
    
    For idxCol = 0 To UBound(aryString)
        wkSheet.Cells(idxRow, idxCol + 1).Value = aryString(idxCol)
    Next
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// �ȉ��A���{���̃R�[���o�b�N
'// ////////////////////////////////////////////////////////////////////////////


'// ////////////////////////////////////////////////////////////////////////////
'// ���\�b�h�F   �V�[�g���j���[ get content
'// �����F       �V�[�g�̃��j���[�\�����s��
'// �����F       control  �ΏۂƂȂ郊�{����̃R���g���[��
'//              content  �߂�l�Ƃ��ĕԂ��A���j���[��\��XML
'// ////////////////////////////////////////////////////////////////////////////
Public Sub sheetMenu_getContent(control As IRibbonControl, ByRef content)
    Dim sheetObj      As Object
    Dim idx           As Integer
    Dim barCtrl_sub   As CommandBarControl
    Dim wkBook        As Workbook
    Dim stMenu        As String
    
    '// $todo:�V�[�g����������ꍇ�̏����ǉ�
    
    Set wkBook = ActiveWorkbook
    idx = 1
    stMenu = "<menu xmlns=""http://schemas.microsoft.com/office/2006/01/customui"" itemSize=""normal"">"
    
    For Each sheetObj In wkBook.Sheets
        If sheetObj.Type = xlWorksheet Then
            '// ID�͐ړ��������ĒʔԂ�ݒ�:MENU_PREFIX + idx
            stMenu = stMenu & "<button id=""" & MENU_PREFIX & CStr(idx) & """ label=""" & sheetObj.Name & """ onAction=""sheetMenuOnAction"""
            If Not sheetObj.Visible Then
                stMenu = stMenu & " enabled=""false"""
            End If
            stMenu = stMenu & " />"
        End If
        idx = idx + 1
    Next
    
    stMenu = stMenu & "</menu>"
    content = stMenu
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// ���\�b�h�F   �V�[�g���A�N�e�B�u�ɂ���
'// �����F       ���j���[�őI�����ꂽ���j���[�L���v�V�������A�A�N�e�B�u���̑Ώۂɂ���
'// �����F       control  �����ꂽ�V�[�g���j���[�B
'// ////////////////////////////////////////////////////////////////////////////
Public Sub sheetMenuOnAction(control As IRibbonControl)
On Error GoTo ErrorHandler
    '// �����ꂽ�V�[�g���j���[��ID�̐ړ���(MENU_PREFIX)�������A�ʔԂ��C���f�b�N�X�Ƃ��Ĉ����ɓn��
    Call ActiveWorkbook.Sheets(CInt(Mid(control.ID, Len(MENU_PREFIX) + 1))).Activate
    Exit Sub

ErrorHandler:
    Call MsgBox(MSG_NO_SHEET, vbOKOnly, APP_TITLE)
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// ���\�b�h�F   ���{���{�^���R�[���o�b�N�Ǘ�
'// �����F       ���{������̃R�[���o�b�N�������ǂ�
'//              �����ꂽ�R���g���[����ID����ɏ������Ăяo���B
'// �����F       control �ΏۃR���g���[��
'// ////////////////////////////////////////////////////////////////////////////
Public Sub ribbonCallback(control As IRibbonControl)
    Select Case control.ID
        '// �V�[�g /////
        Case "SheetComp"                    '// �V�[�g��r
            Call frmCompSheet.Show
        Case "SheetList"                    '// �V�[�g�ꗗ
            Call frmShowSheetList.Show
        Case "SheetSetting"                 '// �V�[�g�̐ݒ�
            Call frmSheetManage.Show
        Case "SheetSortAsc"                 '// �V�[�g�̕��בւ�
            Call psSortWorksheet("ASC")
        Case "SheetSortDesc"                '// �V�[�g�̕��בւ�
            Call psSortWorksheet("DESC")
        
        '// �f�[�^ /////
        Case "Select"                       '// Select�����s
            Call frmGetRecord.Show
        Case "VLookupCopy"                  '// VLookup
            Call psVLookupCopy
        Case "VLookupPaste"
            Call psVLookupPaste
        
        '// �l�̑��� /////
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
        Case "AddLink"                      '// �����N�̒ǉ�
            Call psSetHyperLink
        Case "RemoveLink"                   '// �����N�̍폜
            Call psRemoveHyperLink
        Case "Copy2Clipboard"               '// �Œ蒷�R�s�[
            Call psCopyToClipboard
        Case "Copy2CBMarkdown"               '// �Œ蒷�R�s�[
            Call psCopyToCB_Markdown
            
        '// �r���A�I�u�W�F�N�g /////
        Case "groupRow"                     '// �O���[�v�� �s
            Call psSetGroup_Row
        Case "groupCol"                     '// �O���[�v�� ��
            Call psSetGroup_Col
        Case "removeDup"                    '// �d���̃J�E���g
            Call psDistinctVals
        Case "listDup"                      '// �d�����K�w���ɕ␳
            Call psGroupVals
        
        Case "BorderRowHead"                '// �s�w�b�_�̌r��
            Call gsDrawLine_Header
        Case "BorderColHead"                '// ��w�b�_�̌r��
            Call gsDrawLine_Header_Vert
        Case "BorderData"                   '// �f�[�^�̈�̌r��
            Call gsDrawLine_Data
        Case "FitObjects"                   '// �I�u�W�F�N�g���Z���ɍ��킹��
            Call frmOrderShape.Show
        
        '// �����A�t�@�C�� /////
        Case "AdvancedSearch"               '// �g������
            Call frmSearch.Show
        Case "FileList"                     '// �t�@�C���ꗗ
            Call frmFileList.Show
        
        '// ���̑� /////
        Case "InitTool"                     '// �c�[��������
            Call psInitExTools
        Case "Version"                      '// �o�[�W�������
            Call frmAbout.Show
    End Select

End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// ���\�b�h�F   �V�[�g���N�C�b�N�A�N�Z�X�ɕ\��(Excel2007�ȍ~)
'// �����F       �V�[�g�ꗗ�����j���[�ɕ\������B
'// ////////////////////////////////////////////////////////////////////////////
Public Sub gsShowSheetOnMenu_2007()
    Dim barCtrl       As CommandBar
    
    '// ���j���[�̏�����
    For Each barCtrl In CommandBars
        If barCtrl.Name = "ExSheetMenu" Then
            Call barCtrl.Delete
            Exit For
        End If
    Next
    Set barCtrl = CommandBars.Add(Name:="ExSheetMenu", Position:=msoBarPopup)
    
    Call gsShowSheetOnMenu_sub(barCtrl)
    barCtrl.ShowPopup
    Exit Sub
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// ���\�b�h�F   �V�[�g�����j���[�ɕ\��
'// �����F       �V�[�g�ꗗ�����j���[�ɕ\������B
'// �����F       wkBook: �Ώۃu�b�N
'// ////////////////////////////////////////////////////////////////////////////
Private Sub gsShowSheetOnMenu_sub(barCtrl As Object)
On Error GoTo ErrorHandler
    Const MENU_NUM    As Integer = 30
    
    Dim sheetObj      As Object
    Dim idx           As Integer
    Dim barCtrl_sub   As CommandBarControl
    Dim wkBook        As Workbook
    
    Set wkBook = ActiveWorkbook
    If wkBook.Sheets.Count > MENU_NUM Then
        '// �R�O���ȏ�̃V�[�g�̓O���[�v������
        For Each sheetObj In wkBook.Sheets
            If (sheetObj.Index - 1) Mod MENU_NUM = 0 Then
                Set barCtrl_sub = barCtrl.Controls.Add(Type:=msoControlPopup)
                barCtrl_sub.Caption = "�V�[�g " & CStr(sheetObj.Index) & " �` " & CStr(sheetObj.Index + MENU_NUM - 1) & " (&" & IIf(Int(sheetObj.Index / MENU_NUM) < 10, CStr(Int(sheetObj.Index / MENU_NUM)), Chr(55 + Int(sheetObj.Index / MENU_NUM))) & ")"
            End If
            
            If sheetObj.Type = xlWorksheet Then
                Call psPutMenu(barCtrl_sub.Controls, sheetObj.Name & " (&" & pfGetMenuIndex(sheetObj.Index, MENU_NUM) & ")", "psActivateSheet", IIf(sheetObj.ProtectContents, 505, 0), False, sheetObj.Name, (sheetObj.Visible = xlSheetVisible))
            Else '//If (sheetObj.Type = 4) Or (sheetObj.Type = 1) Then
                Call psPutMenu(barCtrl_sub.Controls, sheetObj.Name & " (&" & pfGetMenuIndex(sheetObj.Index, MENU_NUM) & ")", "psActivateSheet", 422, False, sheetObj.Name, (sheetObj.Visible = xlSheetVisible))
            End If
        Next
    Else
        '// �R�O���ȉ��̃V�[�g�͂��̂܂ܕ\��
        For Each sheetObj In wkBook.Sheets
            If sheetObj.Type = xlWorksheet Then
                Call psPutMenu(barCtrl.Controls, sheetObj.Name & " (&" & IIf(sheetObj.Index < 10, CStr(sheetObj.Index), Chr(55 + sheetObj.Index)) & ")", "psActivateSheet", IIf(sheetObj.ProtectContents, 505, 0), False, sheetObj.Name, (sheetObj.Visible = xlSheetVisible))
            Else '//if (sheetObj.Type = 4) Or (sheetObj.Type = 1) Then
                Call psPutMenu(barCtrl.Controls, sheetObj.Name & " (&" & IIf(sheetObj.Index < 10, CStr(sheetObj.Index), Chr(55 + sheetObj.Index)) & ")", "psActivateSheet", 422, False, sheetObj.Name, (sheetObj.Visible = xlSheetVisible))
            End If
        Next
    End If
    Exit Sub
  
ErrorHandler:
  '// nothing to do.
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// ���\�b�h�F   ���j���[�V���[�g�J�b�g������擾
'// �����F       �V�[�g�̃��j���[�\���ɂāA�V���[�g�J�b�g�p��������擾����
'// �߂�l�F     1�`9�܂���A�`T�̕�����
'// ////////////////////////////////////////////////////////////////////////////
Private Function pfGetMenuIndex(sheetIdx As Integer, menuCnt As Integer) As String
    Select Case sheetIdx Mod menuCnt
        Case 0
            pfGetMenuIndex = Chr(55 + menuCnt)
        Case 1 To 9
            pfGetMenuIndex = CStr(sheetIdx Mod menuCnt)
        Case Else
            pfGetMenuIndex = Chr(55 + (sheetIdx Mod menuCnt))
    End Select
End Function


'// ////////////////////////////////////////////////////////////////////////////
'// ���\�b�h�F   �V�[�g���A�N�e�B�u�ɂ���
'// �����F       ���j���[�őI�����ꂽ���j���[�L���v�V�������A�A�N�e�B�u���̑Ώۂɂ���
'// ////////////////////////////////////////////////////////////////////////////
Public Sub psActivateSheet()
On Error GoTo ErrorHandler
    Call ActiveWorkbook.Sheets(Application.CommandBars.ActionControl.Parameter).Activate
    Exit Sub

ErrorHandler:
    Call MsgBox(MSG_NO_SHEET, vbOKOnly, APP_TITLE)
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// ���\�b�h�F   �A�v���P�[�V�����C�x���g�}��
'// �����F       �e�����O�ɍĕ`���Čv�Z��}�~�ݒ肷��
'// ////////////////////////////////////////////////////////////////////////////
Public Sub gsSuppressAppEvents()
    Application.ScreenUpdating = False                  '// ��ʕ`���~
    Application.Cursor = xlWait                         '// �E�G�C�g�J�[�\��
    Application.EnableEvents = False                    '// �C�x���g�}�~
    Application.Calculation = xlCalculationManual       '// �蓮�v�Z
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// ���\�b�h�F   �A�v���P�[�V�����C�x���g�}������
'// �����F       �e������ɍĕ`���Čv�Z���ĊJ����BgsSuppressAppEvents �̑�
'// ////////////////////////////////////////////////////////////////////////////
Public Sub gsResumeAppEvents()
    Application.StatusBar = False                       '// �X�e�[�^�X�o�[������
    Application.Calculation = xlCalculationAutomatic    '// �����v�Z
    Application.EnableEvents = True
    Application.Cursor = xlDefault
    Application.ScreenUpdating = True
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// END
'// ////////////////////////////////////////////////////////////////////////////
