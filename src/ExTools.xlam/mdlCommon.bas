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


'// ////////////////////////////////////////////////////////////////////////////
'// �A�v���P�[�V�����萔

'// �o�[�W����
Public Const APP_VERSION              As String = "2.2.5.60"                                        '// {���W���[}.{�@�\�C��}.{�o�O�C��}.{�J�����Ǘ��p}

'// �V�X�e���萔
Public Const BLANK                    As String = ""                                                '// �󔒕�����
Public Const DBQ                      As String = """"                                              '// �_�u���N�H�[�g
Public Const CHR_ESC                  As Long = 27                                                  '// Escape �L�[�R�[�h
Public Const CLR_ENABLED              As Long = &H80000005                                          '// �R���g���[���w�i�F �L��
Public Const CLR_DISABLED             As Long = &H8000000F                                          '// �R���g���[���w�i�F ����
Public Const TYPE_RANGE               As String = "Range"                                           '// selection �^�C�v�F�����W
Public Const TYPE_SHAPE               As String = "Shape"                                           '// selection �^�C�v�F�V�F�C�v�ivarType�j
Public Const MENU_PREFIX              As String = "sheet"
Public Const EXCEL_FILE_EXT           As String = "*.xls; *.xlsx"                                   '// �G�N�Z���g���q
Public Const COLOR_ROW                As Integer = 35                                               '// �s�F�����F
Public Const COLOR_DIFF_CELL          As Integer = 3                                                '// �F�F3=��
Public Const COLOR_DIFF_ROW_INS       As Integer = 34                                               '// $mod
Public Const COLOR_DIFF_ROW_DEL       As Integer = 15                                               '// $mod
Public Const EXCEL_PASSWORD           As String = ""                                                '// #017 �G�N�Z�����J���ۂ̃p�X���[�h
Public Const STAT_INTERVAL            As Integer = 100                                              '// �X�e�[�^�X�o�[�X�V�p�x
Public Const ROW_DIFF_STRIKETHROUGH   As Boolean = True                                             '// $mod
Private Const MENU_NUM                As Integer = 30                                               '// �V�[�g�����j���[�ɕ\������ۂ̃O���[�v臒l


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
Public gDatePickerToggle                As Boolean      '// ���t�s�b�J�[�iMonthView�j�\������


'// ////////////////////////////////////////////////////////////////////////////
'// ���\�b�h�F   �c�[��������
'// �����F       ���j���[�̍\���A�A�v���I�u�W�F�N�g�̐ݒ���s���B
'// ////////////////////////////////////////////////////////////////////////////
Private Sub psInitExTools()
    '// ����̐ݒ�
    gLang = Application.LanguageSettings.LanguageID(msoLanguageIDInstall)
    '// �ϐ�������
    gDatePickerToggle = False   '// ���t�s�b�J�[���Ђ炢���܂܂ɂ��遁False
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
    Call gsDrawLine_Data
  
    '// �w�b�_�̏C��
    If headerLines > 0 Then
        Call wkSheet.Range(wkSheet.Cells(1, 1), wkSheet.Cells(headerLines, wkSheet.UsedRange.Columns.Count)).Select
        Call gsDrawLine_Header
    
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
        
        '// �l�̑��� /////
        Case "DatePicker"                       '// ���t
            Call frmDatePicker.Show
        Case "Today", "Now"                     '// ���t - �{�����t/���ݎ���
            Call psPutDateTime(control.ID)
            
        '// �r���A�I�u�W�F�N�g /////
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
'// ���\�b�h�F   ���{���{�^���R�[���o�b�N�Ǘ��i���t�s�b�J�[����g�O���p�j
'// �����F       ���{������̃R�[���o�b�N�������ǂ�
'// �����F       control �ΏۃR���g���[��
'// ////////////////////////////////////////////////////////////////////////////
Public Sub ribbonCallback_DatePickerToggle(control As IRibbonControl, pressed As Boolean)
    gDatePickerToggle = Not pressed
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// ���\�b�h�F   ���{���{�^���R�[���o�b�N�Ǘ��i���t�s�b�J�[����g�O�� ���j���[����p�j
'// �����F       ���{������̃R�[���o�b�N�������ǂ�
'// �����F       control �ΏۃR���g���[��
'// ////////////////////////////////////////////////////////////////////////////
Public Sub GetDatePickerToggleState(control As IRibbonControl, ByRef returnedVal)
    returnedVal = gDatePickerToggle
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
'// ���\�b�h�F   �{�����t/���ݎ����ݒ�
'// �����F       �A�N�e�B�u�Z���ɖ{�����t�܂��͌��ݎ�����ݒ肷��
'// ////////////////////////////////////////////////////////////////////////////
Private Sub psPutDateTime(DateTimeMode As String)
    '// ���O�`�F�b�N�i�A�N�e�B�u�V�[�g�ی�A�I���^�C�v���Z���j
    If Not gfPreCheck(protectCont:=True, selType:=TYPE_RANGE) Then
        Exit Sub
    End If
    
    Call gsSuppressAppEvents
    
    Select Case DateTimeMode
        Case "Today"
            ActiveCell.Value = Date
        Case "Now"
            ActiveCell.Value = Now
    End Select
    
    Call gsResumeAppEvents
    Exit Sub
    
ErrorHandler:
    Call gsResumeAppEvents
    Call gsShowErrorMsgDlg("mdlCommon.psConvValue", Err)
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// END
'// ////////////////////////////////////////////////////////////////////////////
