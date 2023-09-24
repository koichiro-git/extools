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
Public Const APP_VERSION              As String = "3.0.0.77"                                        '// {���W���[}.{�@�\�C��}.{�o�O�C��}.{�J�����Ǘ��p}

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

'Public gADO                             As cADO         '// �ڑ���DB/Excel�I�u�W�F�N�g
Public gLang                            As Long         '// ����


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
Public Sub gsShowErrorMsgDlg(errSource As String, ByVal e As ErrObject)
'Public Sub gsShowErrorMsgDlg(errSource As String, ByVal e As ErrObject, Optional ado As cADO = Nothing)
'    If ado Is Nothing Then
'        '// ADO�I�u�W�F�N�g������̏ꍇ��VB�G���[�Ƃ��Ĉ���
'        Call MsgBox(MSG_ERR & vbLf & vbLf _
'                   & "Error Number: " & e.Number & vbLf _
'                   & "Error Source: " & errSource & vbLf _
'                   & "Error Description: " & e.Description _
'                   , , APP_TITLE)
'        Call e.Clear
'    ElseIf ado.NativeError <> 0 Then
'        '// DB�ł̃G���[�̏ꍇ
'        Call MsgBox(MSG_ERR & vbLf & vbLf _
'                   & "Error Number: " & ado.NativeError & vbLf _
'                   & "Error Source: Database" & vbLf _
'                   & "Error Description: " & ado.ErrorText _
'                   , , APP_TITLE)
'        ado.InitError
'    ElseIf ado.ErrorCode <> 0 Then
'        '// ADO�ł̃G���[�̏ꍇ
'        Call MsgBox(MSG_ERR & vbLf & vbLf _
'                   & "Error Number: " & ado.ErrorCode & vbLf _
'                   & "Error Source: ADO" & vbLf _
'                   & "Error Description�F " & ado.ErrorText _
'                   , , APP_TITLE)
'        ado.InitError
'    Else
        '// ��L�Ŏ�蓦�����ꍇ��VB�G���[�Ƃ��Ĉ���
        Call MsgBox(MSG_ERR & vbLf & vbLf _
                   & "Error Number: " & e.Number & vbLf _
                   & "Error Source: " & errSource & vbLf _
                   & "Error Description: " & e.Description _
                   , , APP_TITLE)
        Call e.Clear
'    End If
End Sub



'// ////////////////////////////////////////////////////////////////////////////
'// ���\�b�h�F   �R���{�{�b�N�X�ݒ�
'// �����F       ������CSV���������ɁA�R���{�{�b�N�X�̒l��ݒ肷��B
'// �����F       targetCombo: �ΏۃR���{�{�b�N�X
'//              propertyStr: �ݒ�l�i{�L�[},{�\��������};{�L�[},{�\��������}...�j
'//              defaultIdx: �����l
'// ////////////////////////////////////////////////////////////////////////////
'Public Sub gsSetCombo(targetCombo As ComboBox, propertyStr As String, defaultIdx As Integer)
'    Dim lineStr()     As String   '// �ݒ�l�̕����񂩂�A�e�s���i�[�i;��؂�j
'    Dim colStr()      As String   '// �e�s�̕����񂩂�A�񂲂Ƃ̒l���i�[�i,��؂�j
'    Dim idxCnt        As Integer
'
'    lineStr = Split(propertyStr, ";")     '//�ݒ�l�̕�������A�s���ɕ���
'
'    Call targetCombo.Clear
'    For idxCnt = 0 To UBound(lineStr)
'        colStr = Split(lineStr(idxCnt), ",")   '//�s�̕�������A�J�������̕�����ɕ���
'        Call targetCombo.AddItem(Trim(colStr(0)))
'        targetCombo.List(idxCnt, 1) = Trim(colStr(1))
'    Next
'
'    targetCombo.ListIndex = defaultIdx    '// �����l��ݒ�
'End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// ���\�b�h�F   �t�H���_�I���_�C�A���O�\��
'// �����F       �t�H���_�I���_�C�A���O��\������B
'// �����F       lngHwnd �E�B���h�E�n���h��
'//              strReturnPath �w�肳�ꂽ�t�H���_�̃p�X������
'// �߂�l�F     True:����  False:���s(�L�����Z����I�������ꍇ�܂�)
'// ////////////////////////////////////////////////////////////////////////////
'Public Function gfShowSelectFolder(ByVal lngHwnd As Long, ByRef strReturnPath) As Boolean
'    Dim lngRet        As Long
'    Dim lngReturnCode As LongPtr
'    Dim strPath       As String
'    Dim biInfo        As BROWSEINFO
'
'    lngRet = False
'
'    '//������̈�̊m��
'    strPath = String(MAX_PATH + 1, Chr(0))
'
'    ' �\���̂̏�����
'    biInfo.hwndOwner = lngHwnd
'    biInfo.lpszTitle = APP_TITLE
'    biInfo.ulFlags = BIF_RETURNONLYFSDIRS
'
'    '// �t�H���_�I���_�C�A���O�̕\��
'    lngReturnCode = apiSHBrowseForFolder(biInfo)
'
'    If lngReturnCode <> 0 Then
'        Call apiSHGetPathFromIDList(lngReturnCode, strPath)
'        strReturnPath = Left(strPath, InStr(strPath, vbNullChar) - 1)
'        gfShowSelectFolder = True
'    Else
'        gfShowSelectFolder = False
'    End If
'End Function


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
    
'    If ActiveSheet Is Nothing Then                              '// �V�[�g�i�u�b�N�j���J����Ă��邩
'        Call MsgBox(MSG_NO_BOOK, vbOKOnly, APP_TITLE)
'        gfPreCheck = False
'        Exit Function
'    End If
'
'    If protectCont And ActiveSheet.ProtectContents Then         '// �A�N�e�B�u�V�[�g���ی삳��Ă��邩
'        Call MsgBox(MSG_SHEET_PROTECTED, vbOKOnly, APP_TITLE)
'        gfPreCheck = False
'        Exit Function
'    End If
'
'    If protectBook And ActiveWorkbook.ProtectStructure Then     '// �u�b�N���ی삳��Ă��邩
'        Call MsgBox(MSG_BOOK_PROTECTED, vbOKOnly, APP_TITLE)
'        gfPreCheck = False
'        Exit Function
'    End If
    
    '// �I��͈͂̃^�C�v���`�F�b�N
'    Select Case selType
'        Case TYPE_RANGE
'            If TypeName(Selection) <> TYPE_RANGE Then
'                Call MsgBox(MSG_NOT_RANGE_SELECT, vbOKOnly, APP_TITLE)
'                gfPreCheck = False
'                Exit Function
'            End If
'        Case TYPE_SHAPE
'            If Not VarType(ActiveWindow.Selection) = vbObject Then
'                Call MsgBox(MSG_SHAPE_NOT_SELECTED, vbOKOnly, APP_TITLE)
'                gfPreCheck = False
'                Exit Function
'            End If
'        Case BLANK
'            '// null
'    End Select
    
'    '// �I��͈̓J�E���g
'    If selAreas > 1 Then
'        If Selection.Areas.Count > selAreas Then
'            Call MsgBox(MSG_TOO_MANY_RANGE, vbOKOnly, APP_TITLE)
'            gfPreCheck = False
'            Exit Function
'        End If
'    End If
'
'    '// �I��͈̓Z���J�E���g
'    If selCols > 1 Then
'        If Selection.Columns.Count > selCols Then
'            Call MsgBox(MSG_TOO_MANY_COLS_8, vbOKOnly, APP_TITLE)
'            gfPreCheck = False
'            Exit Function
'        End If
'    End If
End Function




'// ////////////////////////////////////////////////////////////////////////////
'// ���\�b�h�F   ���{���{�^���R�[���o�b�N�Ǘ�
'// �����F       ���{������̃R�[���o�b�N�������ǂ�
'//              �����ꂽ�R���g���[����ID����ɏ������Ăяo���B
'// �����F       control �ΏۃR���g���[��
'// ////////////////////////////////////////////////////////////////////////////
Public Sub ribbonCallback(control As IRibbonControl)
    Select Case control.ID
'        '// �V�[�g /////
'        Case "SheetComp"                    '// �V�[�g��r
'            Call frmCompSheet.Show
'        Case "SheetList"                    '// �V�[�g�ꗗ
'            Call frmShowSheetList.Show
'        Case "SheetSetting"                 '// �V�[�g�̐ݒ�
'            Call frmSheetManage.Show
'        Case "SheetSortAsc"                 '// �V�[�g�̕��בւ�
'            Call psSortWorksheet("ASC")
'        Case "SheetSortDesc"                '// �V�[�g�̕��בւ�
'            Call psSortWorksheet("DESC")
'
'        '// �f�[�^ /////
'        Case "Select"                       '// Select�����s
'            Call frmGetRecord.Show
'
'        '// �l�̑��� /////
'        Case "DatePicker"                   '// ���t
'            Call frmDatePicker.Show
'        Case "Today", "Now"                 '// ���t - �{�����t/���ݎ���
'            Call psPutDateTime(control.ID)
'
'        '// �r���A�I�u�W�F�N�g /////
'        Case "FitObjects"                   '// �I�u�W�F�N�g���Z���ɍ��킹��
'            Call frmOrderShape.Show
'        Case "AdjShapeAngle"                '// �~�̊p�x��ݒ�
'            Call frmAdjustArch.Show
'        '// �����A�t�@�C�� /////
'        Case "AdvancedSearch"               '// �g������
'            Call frmSearch.Show
'        Case "FileList"                     '// �t�@�C���ꗗ
'            Call frmFileList.Show
'
'        '// ���̑� /////
'        Case "InitTool"                     '// �c�[��������
'            Call psInitExTools
'        Case "Version"                      '// �o�[�W�������
'            Call frmAbout.Show
    End Select

End Sub


''// ////////////////////////////////////////////////////////////////////////////
''// ���\�b�h�F   �A�v���P�[�V�����C�x���g�}��
''// �����F       �e�����O�ɍĕ`���Čv�Z��}�~�ݒ肷��
''// ////////////////////////////////////////////////////////////////////////////
'Public Sub gsSuppressAppEvents()
'    Application.ScreenUpdating = False                  '// ��ʕ`���~
'    Application.Cursor = xlWait                         '// �E�G�C�g�J�[�\��
'    Application.EnableEvents = False                    '// �C�x���g�}�~
'    If Workbooks.Count > 0 Then
'        Application.Calculation = xlCalculationManual       '// �蓮�v�Z
'    End If
'End Sub
'
'
''// ////////////////////////////////////////////////////////////////////////////
''// ���\�b�h�F   �A�v���P�[�V�����C�x���g�}������
''// �����F       �e������ɍĕ`���Čv�Z���ĊJ����BgsSuppressAppEvents �̑�
''// ////////////////////////////////////////////////////////////////////////////
'Public Sub gsResumeAppEvents()
'    Application.StatusBar = False                       '// �X�e�[�^�X�o�[������
'    Application.EnableEvents = True
'    Application.Cursor = xlDefault
'    Application.ScreenUpdating = True
'
'    If Workbooks.Count > 0 Then
'        Application.Calculation = xlCalculationAutomatic    '// �����v�Z
'    End If
'End Sub





'// ////////////////////////////////////////////////////////////////////////////
'// ���\�b�h�F   �V�F�C�v���e�L�X�g�擾
'// �����F       �V�F�C�v���̃e�L�X�g���擾����BCharacters���\�b�h���T�|�[�g���Ȃ��ꍇ�͗�O�����Ńn���h�����O
'//              psExecSearch_Shape�œ��肳�ꂽ�V�F�C�v���̃e�L�X�g��߂�
'//              V3 ����p�u���b�N�֐��Ƃ���frmSearch �� mdlCommon �ֈړ�
'// �����F       shapeObj: �ΏۃV�F�C�v�I�u�W�F�N�g
'// �߂�l�F     �V�F�C�v���̃e�L�X�g�B�V�F�C�v���e�L�X�g���T�|�[�g���Ă��Ȃ��ꍇ�͈ꗥ�Ńu�����N
'// ////////////////////////////////////////////////////////////////////////////
'Public Function gfGetShapeText(shapeObj As Shape) As String
'On Error GoTo ErrorHandler
'    If shapeObj.Type = msoTextEffect Then '// ���[�h�A�[�g�̏ꍇ
'        gfGetShapeText = shapeObj.TextEffect.Text
'    Else
'        gfGetShapeText = shapeObj.TextFrame.Characters.Text
'    End If
'Exit Function
'
'ErrorHandler:
'    gfGetShapeText = BLANK
'End Function


'// ////////////////////////////////////////////////////////////////////////////
'// END
'// ////////////////////////////////////////////////////////////////////////////
