Attribute VB_Name = "mdlLang"
'// ////////////////////////////////////////////////////////////////////////////
'// �v���W�F�N�g   : �g���c�[�� �ǉ��p�b�N
'// �^�C�g��       : ���[�J���C�Y�ݒ�
'// ���W���[��     : mdlLang
'// ����           : �e������L�̐ݒ�
'// ////////////////////////////////////////////////////////////////////////////
'// Copyright (c) by Koichiro.
'// ////////////////////////////////////////////////////////////////////////////
Option Explicit

'// ����R�[�h Application.LanguageSettings.LanguageID(msoLanguageIDInstall) �œ�����l
#Const cLANG = 1041   '// ���{��
'#Const cLANG = 1033  '// English


'// ////////////////////////////////////////////////////////////////////////////
'// 1041 - Japanese
#If cLANG = 1041 Then


'// ////////////////////
'// �A�v�����ʕϐ� (�ϐ�����: APP_{string} )
Public Const APP_TITLE                          As String = "�g���c�[��"
Public Const APP_EXL_FILE                       As String = "�G�N�Z���`�� �t�@�C�� (#),#"


'// ////////////////////
'// ���j���[ (�ϐ�����: MENU_{string} )


'// ////////////////////
'// ���b�Z�[�W (�ϐ�����: MSG_{string} )

'// ����
Public Const MSG_ERR                            As String = "�������ɃG���[���������܂����B"
Public Const MSG_NO_BOOK                        As String = "�u�b�N������܂���B"
Public Const MSG_FINISHED                       As String = "�������I�����܂����B"
Public Const MSG_PROCEED_CREATE_SHEET           As String = "�V�[�g�쐬�́A���݂̃u�b�N�ɃV�[�g��ǉ����܂��B��낵���ł����H"
Public Const MSG_TOO_MANY_RANGE                 As String = "���̃R�}���h�͕����̑I��͈͂ɑ΂��Ď��s�ł��܂���B"
Public Const MSG_NOT_NUMERIC                    As String = "���l����͂��Ă��������B"
Public Const MSG_ZERO_NOT_ACCEPTED              As String = "�[���͓��͂ł��܂���B"
Public Const MSG_NO_DIR                         As String = "�����p�X���w�肳��Ă��܂���B"
Public Const MSG_DIR_NOT_EXIST                  As String = "�w�肳�ꂽ�����p�X�͑��݂��܂���B"
Public Const MSG_NO_RESULT                      As String = "�������ʂ̓[�����ł��B"
Public Const MSG_SHEET_PROTECTED                As String = "�V�[�g���ی삳��Ă��܂��B�V�[�g�̕ی���������ĉ������B"
Public Const MSG_BOOK_PROTECTED                 As String = "�u�b�N���ی삳��Ă��܂��B�u�b�N�̕ی���������ĉ������B"
Public Const MSG_INVALID_NUM                    As String = "�L���Ȑ��l��ݒ肵�Ă��������B"
Public Const MSG_CONFIRM                        As String = "���s���܂��B��낵���ł����H"
Public Const MSG_NO_SHEET                       As String = "�V�[�g��������܂���B"
Public Const MSG_INVALID_RANGE                  As String = "�L���Ȓl�̂���͈͂�I�����Ă��������B"
Public Const MSG_TOO_MANY_COLS_8                As String = "�W��ȏ�̑I��͈͂͏����ł��܂���B"
Public Const MSG_NOT_RANGE_SELECT               As String = "�Z����I�����ĉ������B"
Public Const MSG_VLOOKUP_MASTER_2COLS           As String = "VLOOKUP�̃}�X�^�\�Ƃ���2��ȏ��I�����ĉ������B"
Public Const MSG_VLOOKUP_SET_2COLS              As String = "VLOOKUP�̓\��t����Ƃ���2��ȏ��I�����ĉ������B"
Public Const MSG_VLOOKUP_SEL_DUPLICATED         As String = "VLOOKUP�̃}�X�^�\�Ɠ\��t���悪�d�����Ă��܂��B"
Public Const MSG_VLOOKUP_NO_MASTER              As String = "VLOOKUP�̃}�X�^�\���I������Ă��܂���B"
Public Const MSG_SEL_DEFAULT_COLOR              As String = "�I�����ꂽ�Z���̓f�t�H���g�F���w�肳��Ă��܂��B�����𑱂��܂����H"
Public Const MSG_SHAPE_NOT_SELECTED             As String = "�V�F�C�v���I������Ă��܂���B"
Public Const MSG_PROCESSING                     As String = "���s���ł�..."
Public Const MSG_SHAPE_MULTI_SELECT             As String = "2�ȏ�̃V�F�C�v��I�����Ă��������B"
Public Const MSG_BARCODE_NOT_AVAILABLE          As String = "�o�[�R�[�h��Excel2016�ȍ~�Ŏg�p�\�ł��B"

'// ���b�Z�[�W�FfrmTranslation
Public Const MSG_SERVICE_TRANS_NOT_REACHABLE    As String = "�|��T�C�g�ɃA�N�Z�X�ł��܂���Bini�t�@�C���̐ݒ�ƃC���^�[�l�b�g�ڑ����m�F���Ă��������B"


'// ////////////////////
'// �t�H�[�����x�� (�ϐ�����: LBL_{form code}_{string} )

'// ����
Public Const LBL_COM_EXEC                       As String = "���s"
Public Const LBL_COM_CLOSE                      As String = "����"
Public Const LBL_COM_BROWSE                     As String = "�Q��..."
Public Const LBL_COM_TARGET                     As String = "�o�͑Ώ�"
Public Const LBL_COM_OPTIONS                    As String = "�o�̓I�v�V����"
Public Const LBL_COM_CHAR_SET                   As String = "�����R�[�h"
Public Const LBL_COM_CR_CODE                    As String = "���s�R�[�h"
Public Const LBL_COM_NEW_SHEET                  As String = "�V�[�g�쐬"
Public Const LBL_COM_CHECK_ALL                  As String = "���ׂđI��"
Public Const LBL_COM_UNCHECK                    As String = "�I������"
Public Const LBL_COM_HYPERLINK                  As String = "�n�C�p�[�����N�̐ݒ�"

'// frmTranslation (TRN)
Public Const LBL_TRN_FORM                       As String = "�|��"
Public Const LBL_TRN_KEY                        As String = "�F�؃L�["
Public Const LBL_TRN_LANG                       As String = "�|�󌾌�"
Public Const LBL_TRN_OUTPUT                     As String = "�o�͌`��"


'// ////////////////////
'// �R���{�{�b�N�X (�ϐ�����: CMB_{form code}_{string} )

'// ����
Public Const CMB_COM_CHAR_SET                   As String = "0,S-JIS;1,JIS;2,EUC;3,Unicode(UTF-8)"
Public Const CMB_COM_CR_CODE                    As String = "0,CR(#13) + LF(#10);1,LF(#10);2,CR(#13)"

'// frmTranslation (TRN)
Public Const CMB_TRN_LANGUAGE                   As String = "ja,�p�� �� ���{��;en,���{�� �� �p��"
Public Const CMB_TRN_OUTPUT                     As String = "0,�����̉��ɒǉ�;1,�����̏�ɒǉ�;2,�������폜���ď㏑��;3,�E�̃Z���ɏ㏑��;4,�R�����g�ɒǉ�"


'// ////////////////////
'// �ꗗ�o�̓w�b�_



'// ////////////////////////////////////////////////////////////////////////////
'// 1033 - English
#ElseIf cLANG = 1033 Then

'// ////////////////////
'// �A�v�����ʕϐ� (�ϐ�����: APP_{string} )
Public Const APP_TITLE                          As String = "Excel Extentions"
Public Const APP_EXL_FILE                       As String = "Excel file (#),#"


'// ////////////////////
'// ���j���[ (�ϐ�����: MENU_{string} )


'// ////////////////////
'// ���b�Z�[�W (�ϐ�����: MSG_{string} )

'// ����
Public Const MSG_ERR                            As String = "An error occured during the operation."
Public Const MSG_NO_BOOK                        As String = "No books opened."
Public Const MSG_FINISHED                       As String = "Operation finished successfully."
Public Const MSG_PROCEED_CREATE_SHEET           As String = "Are you sure to append new sheet on the current workbook?"
Public Const MSG_TOO_MANY_RANGE                 As String = "This operation supports single selection area only."
Public Const MSG_NOT_NUMERIC                    As String = "Valid number is required."
Public Const MSG_ZERO_NOT_ACCEPTED              As String = "Zero is not supported."
Public Const MSG_NO_DIR                         As String = "Please identify the search path."
Public Const MSG_DIR_NOT_EXIST                  As String = "The search path is not correct."
Public Const MSG_NO_RESULT                      As String = "No result found."
Public Const MSG_SHEET_PROTECTED                As String = "Current sheet is protected.  Please unprotect and execute again."
Public Const MSG_BOOK_PROTECTED                 As String = "Current book is protected.  Please unprotect and execute again."
Public Const MSG_INVALID_NUM                    As String = "Valid number is required."
Public Const MSG_CONFIRM                        As String = "Do you want to proceed?"
Public Const MSG_NO_SHEET                       As String = "No sheet found."
Public Const MSG_INVALID_RANGE                  As String = "Selected area is not correct."
Public Const MSG_TOO_MANY_COLS_8                As String = "Target columns should be less than 9 columns."
Public Const MSG_NOT_RANGE_SELECT               As String = "Selected area is not correct."
Public Const MSG_VLOOKUP_MASTER_2COLS           As String = "VLOOKUP master table needs two or more columns."
Public Const MSG_VLOOKUP_SET_2COLS              As String = "VLOOKUP paste area needs two or more columns."
Public Const MSG_VLOOKUP_SEL_DUPLICATED         As String = "Cannot paste VLOOKUP on master table area."
Public Const MSG_VLOOKUP_NO_MASTER              As String = "Nothing to paste.  Please select VLOOKUP master table first."
Public Const MSG_SEL_DEFAULT_COLOR              As String = "Default color is set on the current cell.  Do you want to proceed?"
Public Const MSG_SHAPE_NOT_SELECTED             As String = "Shapes not selected."
Public Const MSG_PROCESSING                     As String = "Processing..."
Public Const MSG_SHAPE_MULTI_SELECT             As String = "Please select 2 or more target shapes."
Public Const MSG_BARCODE_NOT_AVAILABLE          As String = "Barcode function is available on Excel 2016 or higher."

'// ���b�Z�[�W�FfrmTranslation
Public Const MSG_SERVICE_TRANS_NOT_REACHABLE    As String = "Translation service unreachable. Please check your ini file settings and internet connection."


'// ////////////////////
'// �t�H�[�����x�� (�ϐ�����: LBL_{form code}_{string} )

'// ����
Public Const LBL_COM_EXEC                       As String = "Execute"
Public Const LBL_COM_CLOSE                      As String = "Close"
Public Const LBL_COM_BROWSE                     As String = "Browse"
Public Const LBL_COM_TARGET                     As String = "Target sheets"
Public Const LBL_COM_OPTIONS                    As String = "Options"
Public Const LBL_COM_CHAR_SET                   As String = "Character set"
Public Const LBL_COM_CR_CODE                    As String = "CR code"
Public Const LBL_COM_NEW_SHEET                  As String = "New sheet"
Public Const LBL_COM_CHECK_ALL                  As String = "Check all"
Public Const LBL_COM_UNCHECK                    As String = "Uncheck all"
Public Const LBL_COM_HYPERLINK                  As String = "Set hyperlinks"

'// frmTranslation (TRN)
Public Const LBL_TRN_FORM                       As String = "Translate"
Public Const LBL_TRN_KEY                        As String = "Auth Key"
Public Const LBL_TRN_LANG                       As String = "Target Language"
Public Const LBL_TRN_OUTPUT                     As String = "Output"


'// ////////////////////
'// �R���{�{�b�N�X (�ϐ�����: CMB_{form code}_{string} )

'// ����
Public Const CMB_COM_CHAR_SET                   As String = "0,S-JIS;1,JIS;2,EUC;3,Unicode(UTF-8)"
Public Const CMB_COM_CR_CODE                    As String = "0,CR(#13) + LF(#10);1,LF(#10);2,CR(#13)"

'// frmTranslation (TRN)
Public Const CMB_TRN_LANGUAGE                   As String = "ja,English �� Japanese;en,Japanese �� English"
Public Const CMB_TRN_OUTPUT                     As String = "0,Below the original;1,Top of the original;2,Overwrite;3,Right cell;4,Comment"


'// ////////////////////
'// �ꗗ�o�̓w�b�_


#End If


'// ////////////////////////////////////////////////////////////////////////////
'// END
'// ////////////////////////////////////////////////////////////////////////////
