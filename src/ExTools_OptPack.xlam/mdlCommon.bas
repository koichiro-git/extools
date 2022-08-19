Attribute VB_Name = "mdlCommon"
'// ////////////////////////////////////////////////////////////////////////////
'// �v���W�F�N�g   : �g���c�[�� �ǉ��p�b�N
'// �^�C�g��       : ���ʊ֐�
'// ���W���[��     : mdlCommon
'// ����           : �V�X�e���̋��ʊ֐��A�N�����̐ݒ�Ȃǂ��Ǘ�
'// ////////////////////////////////////////////////////////////////////////////
'// Copyright (c) by Koichiro.
'// ////////////////////////////////////////////////////////////////////////////
Option Explicit
Option Base 0

'// ////////////////////////////////////////////////////////////////////////////
'// �A�v���P�[�V�����萔

'// �o�[�W����
Public Const OPTION_PACK_VERSION      As String = "2"                                               '// ���̃��W���[���ŗL�̃o�[�W�����i�Ǘ��p�ʂ��ԍ��j

'// �V�X�e���萔
Public Const PROJECT_NAME             As String = "ExToolsOptionalPack"                             '// �{�A�h�C������
Public Const BLANK                    As String = ""                                                '// �󔒕�����
Public Const CHR_ESC                  As Long = 27                                                  '// Escape �L�[�R�[�h
Public Const CLR_ENABLED              As Long = &H80000005                                          '// �R���g���[���w�i�F �L��
Public Const CLR_DISABLED             As Long = &H8000000F                                          '// �R���g���[���w�i�F ����
Public Const TYPE_RANGE               As String = "Range"                                           '// selection �^�C�v�F�����W
Public Const TYPE_SHAPE               As String = "Shape"                                           '// selection �^�C�v�F�V�F�C�v�ivarType�j
Public Const MENU_PREFIX              As String = "sheet"


'// ////////////////////////////////////////////////////////////////////////////
'// Windows API �֘A�̐錾

'// ini�t�@�C���ǂݍ���
Declare PtrSafe Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long


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
'// ���\�b�h�F   �A�v���P�[�V�����C�x���g�}��
'// �����F       �e�����O�ɍĕ`���Čv�Z��}�~�ݒ肷��
'// ////////////////////////////////////////////////////////////////////////////
Public Sub gsSuppressAppEvents()
    Application.ScreenUpdating = False                  '// ��ʕ`���~
    Application.Cursor = xlWait                         '// �����v�J�[�\��
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
'// ���\�b�h�F   �G���[���b�Z�[�W�\���iVBA�j
'// �����F       ��O�������ŏ����ł��Ȃ���O�̃G���[�̓��e���A�_�C�A���O�\������B
'// �����F       errSource: �G���[�̔������̃I�u�W�F�N�g�܂��̓A�v���P�[�V�����̖��O������������
'//              e: �u�a�G���[�I�u�W�F�N�g
'//              objAdo�F ADO�I�u�W�F�N�g�i�ȗ��j
'// ////////////////////////////////////////////////////////////////////////////
Public Sub gsShowErrorMsgDlg_VBA(errSource As String, ByVal e As ErrObject)
    Call MsgBox(MSG_ERR & vbLf & vbLf _
               & "Error Number: " & e.Number & vbLf _
               & "Error Source: " & errSource & vbLf _
               & "Error Description: " & e.Description _
               , , APP_TITLE)
    Call e.Clear
End Sub


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
'// ���\�b�h�F   ���{���{�^���R�[���o�b�N�Ǘ�
'// �����F       ���{������̃R�[���o�b�N�������ǂ�
'//              �����ꂽ�R���g���[����ID����ɏ������Ăяo���B
'// �����F       control �ΏۃR���g���[��
'// ////////////////////////////////////////////////////////////////////////////
Public Sub ribbonCallback2(control As IRibbonControl)
    Select Case control.ID
        Case "FormatPhoneNumbers"                       '// �d�b�ԍ��␳
            Call gsFormatPhoneNumbers
        Case "Translation"                              '// �|��
            Call frmTranslation.Show
    End Select
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// ���\�b�h�F   ini�t�@�C���ݒ�l�擾
'// �����F       xlam�t�@�C���Ɠ�����ini�t�@�C������w�肳�ꂽ�l���擾����
'// �����F       section �Z�N�V����
'//              key     ���ʃL�[
'// ////////////////////////////////////////////////////////////////////////////
Public Function gfGetIniFileSetting(section As String, key As String) As String
    Dim sValue      As String   '// �擾�o�b�t�@
    Dim lSize       As Long     '// �擾�o�b�t�@�̃T�C�Y
    Dim lRet        As Long     '// �߂�l
    
    '// �擾�o�b�t�@������
    lSize = 100
    sValue = Space(lSize)
    
    lRet = GetPrivateProfileString(section, key, BLANK, sValue, lSize, Replace(Application.VBE.VBProjects(PROJECT_NAME).Filename, ".xlam", ".ini"))
    gfGetIniFileSetting = Trim(Left(sValue, InStr(sValue, Chr(0)) - 1))
End Function


'// ////////////////////////////////////////////////////////////////////////////
'// END
'// ////////////////////////////////////////////////////////////////////////////
