Attribute VB_Name = "mdlLang"
'// ////////////////////////////////////////////////////////////////////////////
'// �v���W�F�N�g   : �g���c�[��
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
'Public Const APP_SQL_FILE                       As String = "SQL�t�@�C�� (*.sql; *.txt),*.sql;*.txt"
Public Const APP_EXL_FILE                       As String = "�G�N�Z���`�� �t�@�C�� (#),#"
'Public Const APP_XML_FILE                       As String = "XML�t�@�C�� (*.xml),*.xml"


'// ////////////////////
'// ���j���[ (�ϐ�����: MENU_{string} )
'Public Const MENU_SHEET_MENU                    As String = "�V�[�g(&S)"
'Public Const MENU_EXTOOL                        As String = "�g��(&X)"
'Public Const MENU_SHEET_GROUP                   As String = "�V�[�g # �` @"
'Public Const MENU_COMP_SHEET                    As String = "�V�[�g/�u�b�N��r..."
'Public Const MENU_SELECT                        As String = "SQL�����s..."
''Public Const MENU_FILE_EXP                      As String = "DML/�f�[�^�o��..."
'Public Const MENU_SORT                          As String = "�V�[�g�̃\�[�g"
'Public Const MENU_SORT_ASC                      As String = "�����\�[�g"
'Public Const MENU_SORT_DESC                     As String = "�~���\�[�g"
'Public Const MENU_SHEET_LIST                    As String = "�V�[�g�ꗗ���o��..."
'Public Const MENU_SHEET_SETTING                 As String = "�V�[�g�̐ݒ�..."
Public Const MENU_CHANGE_CHAR                   As String = "������̕ϊ�"
Public Const MENU_CAPITAL                       As String = "�啶��"
Public Const MENU_SMALL                         As String = "������"
Public Const MENU_PROPER                        As String = "�P��̐擪������啶��"
Public Const MENU_ZEN                           As String = "�S�p"
Public Const MENU_HAN                           As String = "���p"
Public Const MENU_TRIM                          As String = "�l�̃g����"
'Public Const MENU_SELECTION                     As String = "�I��͈͂̐ݒ�"
'Public Const MENU_SELECTION_UNIT                As String = "#�s��"
'Public Const MENU_CLIPBOARD                     As String = "�N���b�v�{�[�h�փR�s�[(&C)"
'Public Const MENU_DRAW_LINE_H                   As String = "�w�b�_���̌r����`��"
'Public Const MENU_DRAW_LINE_D                   As String = "�f�[�^���̌r����`��"
'Public Const MENU_RESET                         As String = "�g���c�[���̏�����"
'Public Const MENU_VERSION                       As String = "�o�[�W�������"
'Public Const MENU_LINK                          As String = "�n�C�p�[�����N"
'Public Const MENU_LINK_ADD                      As String = "�n�C�p�[�����N�̐ݒ�"
'Public Const MENU_LINK_REMOVE                   As String = "�n�C�p�[�����N�̍폜"
'Public Const MENU_DRAW_LINE_H_HORIZ             As String = "�\�̏㕔�i�����j"
'Public Const MENU_DRAW_LINE_H_VERT              As String = "�\�̍��i�c���j"
'Public Const MENU_XML                           As String = "XML����..."
'Public Const MENU_FILE                          As String = "�t�@�C���ꗗ�o��..."
'Public Const MENU_GROUP                         As String = "�O���[�v����"
'Public Const MENU_GROUP_SET_ROW                 As String = "�O���[�v���i�s�j"
'Public Const MENU_GROUP_SET_COL                 As String = "�O���[�v���i��j"
'Public Const MENU_GROUP_DISTINCT                As String = "�d���f�[�^�̏W��"
'Public Const MENU_GROUP_VALUE                   As String = "�d���f�[�^���K�w���ɕ␳"
'Public Const MENU_DRAW_CHART                    As String = "�ȈՃ`���[�g��`��..."
'Public Const MENU_SEARCH                        As String = "�g������(&S)..."
'Public Const MENU_RESIZE_SHAPE                  As String = "�V�F�C�v���Z���ɍ��킹��..."
'Public Const MENU_SELECTION_BACK_COLR           As String = "�����w�i�F"
'Public Const MENU_SELECTION_FONT_COLR           As String = "�����t�H���g�F"


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

'// ���b�Z�[�W�FfrmCompSheet
Public Const MSG_ERROR_NEED_BOOKNAME            As String = "�u�b�N���w�肵�Ă��������B"
Public Const MSG_NO_FILE                        As String = "�t�@�C�����J���܂���"
Public Const MSG_UNMATCH_SHEET                  As String = "�V�[�g�\�����قȂ�܂��B"
Public Const MSG_NO_DIFF                        As String = "�f�[�^�͓����ł��B"
Public Const MSG_SHEET_NAME                     As String = "�V�[�g�����قȂ�܂�"
Public Const MSG_INS_ROW                        As String = "�s�ǉ�"
Public Const MSG_DEL_ROW                        As String = "�s�폜"

'// ���b�Z�[�W�FfrmSheetManage
Public Const MSG_VAL_10_400                     As String = "�Y�[���l�ɗL���Ȑ��l����͂��ĉ������B(10�`400)"
Public Const MSG_SHEETS_PROTECTED               As String = "�ی삳��Ă���V�[�g���P�ȏ゠��܂��B�V�[�g�̕ی���������ĉ������B"
Public Const MSG_COMPLETED_FILES                As String = "�ȉ��̃t�@�C���ɂ��Ă̏����͏I�����Ă��܂��B�O�̂��߃t�@�C���̍X�V���t���m�F���ĉ������B"

'// ���b�Z�[�W�FfrmGetRecord
Public Const MSG_TOO_MANY_ROWS                  As String = "�o�͉\�ȍő匏���ɒB���܂����B�ȍ~�̃f�[�^�͐؂�̂Ă��܂��B"
Public Const MSG_TOO_MANY_COLS                  As String = "�񐔂������l���z���Ă��܂��B�������z������͐؂�̂Ă��܂��B"
Public Const MSG_QUERY                          As String = "�₢���킹��"
Public Const MSG_EXTRACT_SHEET                  As String = "�V�[�g�֏o�͒�"
Public Const MSG_PAGE_SETUP                     As String = "�����ݒ蒆"
Public Const MSG_ROWS_PROCESSED                 As String = "�s����������܂���"

'// ���b�Z�[�W�FfrmDataExport
'Public Const MSG_TABLE_NAME                     As String = "�\�_������"
'Public Const MSG_COLUMN_NAME                    As String = "�񖼏�"

'// ���b�Z�[�W�FfrmDrawChart
'Public Const MSG_INVALID_COL_MIN                As String = "�O���t��A���荶�ɂ͕`��ł��܂���B"
'Public Const MSG_INVALID_COL_MAX                As String = "�O���t�͍ő����E�ɂ͕`��ł��܂���B"

'// ���b�Z�[�W�FfrmLogin
Public Const MSG_LOG_ON_SUCCESS                 As String = "���O�C�����܂����B"
Public Const MSG_LOG_ON_FAILED                  As String = "���O�C���Ɏ��s���܂����B"
Public Const MSG_NEED_FILL_ID                   As String = "���[�U�h�c�����͂���Ă��܂���B"
Public Const MSG_NEED_FILL_PWD                  As String = "�p�X���[�h�����͂���Ă��܂���B"
Public Const MSG_NEED_FILL_TNS                  As String = "�ڑ������񂪓��͂���Ă��܂���B"
Public Const MSG_NEED_EXCEL_SAVED               As String = "���݂̃u�b�N�͕ۑ�����Ă��܂���B�u�b�N��ۑ����Ă��������B"

'// ���b�Z�[�W�FfrmSearch
Public Const MSG_NO_CONDITION                   As String = "�����������w�肵�ĉ������B"
Public Const MSG_WRONG_COND                     As String = "���������Ɏw�肳�ꂽ������͖����ł��B"

'// ���b�Z�[�W�FfrmFileList
Public Const MSG_MAX_DEPTH                      As String = "�ő�[�x�ɒB���܂����B"
Public Const MSG_ERR_PRIV                       As String = "�G���[�F�A�N�Z�X���Ȃǂɖ�肪����\��������܂��B"
Public Const MSG_EMPTY_DIR                      As String = "��f�B���N�g��"
Public Const MSG_ZERO_BYTE                      As String = "�[���o�C�g"



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

'// frmCompSheet (CMP)
Public Const LBL_CMP_FORM                       As String = "�V�[�g/�u�b�N��r"
Public Const LBL_CMP_MODE_SHEET                 As String = "�V�[�g��r"
Public Const LBL_CMP_MODE_BOOK                  As String = "�u�b�N��r"
Public Const LBL_CMP_SHEET1                     As String = "��r���V�[�g"
Public Const LBL_CMP_SHEET2                     As String = "��r��V�[�g"
Public Const LBL_CMP_BOOK1                      As String = "��r���u�b�N"
Public Const LBL_CMP_BOOK2                      As String = "��r��u�b�N"
Public Const LBL_CMP_OPTIONS                    As String = "��r�I�v�V����"
Public Const LBL_CMP_RESULT                     As String = "�o�͐�"
Public Const LBL_CMP_MARKER                     As String = "�}�[�J�["
Public Const LBL_CMP_METHOD                     As String = "��r���@"
Public Const LBL_CMP_SHOW_COMMENT               As String = "�ύX�ӏ��̃R�����g��\��"

'// frmShowSheetList (SSL)
Public Const LBL_SSL_FORM                       As String = "�V�[�g�ꗗ�o��"
Public Const LBL_SSL_TARGET                     As String = "�o�͐�"
Public Const LBL_SSL_OPTIONS                    As String = "�V�[�g�̒l�̏o��"
Public Const LBL_SSL_ROWS                       As String = "�s��"
Public Const LBL_SSL_COLS                       As String = "��"

'// frmSheetManage (SMG)
Public Const LBL_SMG_FORM                       As String = "�V�[�g����"
Public Const LBL_SMG_TARGET                     As String = "�����Ώ�"
Public Const LBL_SMG_SCROLL                     As String = "�X�N���[����������"
Public Const LBL_SMG_FONT_COLOR                 As String = "�t�H���g�F��������"
Public Const LBL_SMG_HYPERLINK                  As String = "�n�C�p�[�����N���폜"
Public Const LBL_SMG_COMMENT                    As String = "�R�����g���폜"
Public Const LBL_SMG_HEAD_FOOT                  As String = "�w�b�_�ƃt�b�^�̕\���ݒ�"
Public Const LBL_SMG_MARGIN                     As String = "�}�[�W����ݒ�"
Public Const LBL_SMG_PAGEBREAK                  As String = "���y�[�W�ƈ���͈͂��N���A"
Public Const LBL_SMG_PRINT_OPT                  As String = "����̊g��/�k��"
Public Const LBL_SMG_PRINT_NONE                 As String = "�ݒ�Ȃ�"
Public Const LBL_SMG_PRINT_100                  As String = "100%"
Public Const LBL_SMG_PRINT_HRZ                  As String = "���P��/�c��"
Public Const LBL_SMG_PRINT_1_PAGE               As String = "���P��/�c�P��"
Public Const LBL_SMG_VIEW                       As String = "�r���["
Public Const LBL_SMG_ZOOM                       As String = "�Y�[��(%)"
Public Const LBL_SMG_AUTOFILTER                 As String = "�I�[�g�t�B���^"

'// frmGetRecord (GRC)
Public Const LBL_GRC_FORM                       As String = "SQL�����s"
Public Const LBL_GRC_FILE                       As String = "�t�@�C��"
Public Const LBL_GRC_OPTIONS                    As String = "�o�̓I�v�V����"
Public Const LBL_GRC_DATE_FORMAT                As String = "���t����"
Public Const LBL_GRC_HEADER                     As String = "�w�b�_�o��"
Public Const LBL_GRC_GROUPING                   As String = "�O���[�v��"
Public Const LBL_GRC_BORDERS                    As String = "�g����\��"
Public Const LBL_GRC_BG_COLOR                   As String = "�s��h�蕪��"
Public Const LBL_GRC_SCRIPT                     As String = "SQL�X�N���v�g"
Public Const LBL_GRC_LOGIN                      As String = "���O�C��"
Public Const LBL_GRC_FILE_OPEN                  As String = "�t�@�C�����J��"
Public Const LBL_GRC_SEARCH                     As String = "���s"

''// frmDataExport (EXP)
'Public Const LBL_EXP_FORM                       As String = "DML/�f�[�^�o��"
'Public Const LBL_EXP_FILE_TYPE                  As String = "�o�͌`��"
'Public Const LBL_EXP_TARGET                     As String = "�o�͑Ώ�"
'Public Const LBL_EXP_OPTIONS                    As String = "�o�̓I�v�V����"
'Public Const LBL_EXP_DATE_FORMAT                As String = "���t����"
'Public Const LBL_EXP_QUOTE                      As String = "�N�H�[�g"
'Public Const LBL_EXP_SEPARATOR                  As String = "��؂蕶��"
'Public Const LBL_EXP_CHAR_SET                   As String = "�����R�[�h"
'Public Const LBL_EXP_CR_CODE                    As String = "���s�R�[�h"
'Public Const LBL_EXP_QUOTE_ALL                  As String = "���l�E���t���N�H�[�g"
'Public Const LBL_EXP_FORMAT_DML                 As String = "DML�����s�Ő��`"
'Public Const LBL_EXP_HEADER                     As String = "�w�b�_�E�t�b�^���o��"
'Public Const LBL_EXP_COL_NAME                   As String = "���ږ����o��"
'Public Const LBL_EXP_SEMICOLON                  As String = "�Z�~�R�������o�͂��Ȃ�"
'Public Const LBL_EXP_NUM_POINT                  As String = "���l�̏����_���o�͂��Ȃ�"
'Public Const LBL_EXP_CREATE_SHEET               As String = "�V�[�g�쐬"
'
''// frmXmlManage (XML)
'Public Const LBL_XML_FORM                       As String = "XML����"
'Public Const LBL_XML_INDENT                     As String = "�C���f���g"
'Public Const LBL_XML_PUT_DEF                    As String = "XML�錾�̏o��"
'Public Const LBL_XML_LOAD                       As String = "�Ǎ�"
'Public Const LBL_XML_WRITE                      As String = "�o��"
'
''// frmDrawChart (CHT)
'Public Const LBL_CHT_FORM                       As String = "�ȈՃ`���[�g�̕`��"
'Public Const LBL_CHT_MAX_VAL                    As String = "�ő�l"
'Public Const LBL_CHT_INTERVAL                   As String = "�⏕���Ԋu"
'Public Const LBL_CHT_POSITION                   As String = "�`��ʒu"
'Public Const LBL_CHT_DIRECTION                  As String = "����"
'Public Const LBL_CHT_GRADATION                  As String = "�O���f�[�V����"
'Public Const LBL_CHT_LEGEND                     As String = "�}��̕\��"
'Public Const LBL_CHT_LINE_FRONT                 As String = "�⏕������O�ɕ\��"

'// frmOrderShape (ORD)
Public Const LBL_ORD_FORM                       As String = "�V�F�C�v�̔z�u"
Public Const LBL_ORD_MARGIN                     As String = "�}�[�W��"
Public Const LBL_ORD_OPTIONS                    As String = "�ڍאݒ�"
Public Const LBL_ORD_HEIGHT                     As String = "�㉺���̐ݒ�"
Public Const LBL_ORD_WIDTH                      As String = "���E���̐ݒ�"

'// frmSearch (SRC)
Public Const LBL_SRC_FORM                       As String = "�g������"
Public Const LBL_SRC_STRING                     As String = "�������镶����"
Public Const LBL_SRC_TARGET                     As String = "�����Ώ�"
Public Const LBL_SRC_MARK                       As String = "�}�[�J�[�̕\��"
Public Const LBL_SRC_DIR                        As String = "��������t�H���_"
Public Const LBL_SRC_SUB_DIR                    As String = "�T�u�t�H���_������"
Public Const LBL_SRC_IGNORE_CASE                As String = "�啶������������ʂ��Ȃ�"
Public Const LBL_SRC_OBJECT                     As String = "��������I�u�W�F�N�g"
Public Const LBL_SRC_CELL_TEXT                  As String = "�Z���̕����������"
Public Const LBL_SRC_CELL_FORMULA               As String = "�Z���̐���������"
Public Const LBL_SRC_SHAPE                      As String = "�V�F�C�v������"
Public Const LBL_SRC_COMMENT                    As String = "�R�����g������"
Public Const LBL_SRC_CELL_NAME                  As String = "�Z�����̂�����"
Public Const LBL_SRC_SHEET_NAME                 As String = "�V�[�g��������"
Public Const LBL_SRC_HYPERLINK                  As String = "�n�C�p�[�����N������"
Public Const LBL_SRC_HEADER                     As String = "�w�b�_�E�t�b�^������"
Public Const LBL_SRC_GRAPH                      As String = "�O���t������"

'// frmFileList (LST)
Public Const LBL_LST_FORM                       As String = "�t�@�C���ꗗ�o��"
Public Const LBL_LST_ROOT                       As String = "�o�̓��[�g"
Public Const LBL_LST_DEPTH                      As String = "�ő�[�x"
Public Const LBL_LST_TARGET                     As String = "�Ώۃt�@�C��"
Public Const LBL_LST_EXT                        As String = "�g���q"
Public Const LBL_LST_SIZE                       As String = "�T�C�Y�P��"
Public Const LBL_LST_REL_PATH                   As String = "���΃p�X�ŕ\��"

'// frmLogin (LGI)
Public Const LBL_LGI_FORM                       As String = "Login"
Public Const LBL_LGI_UID                        As String = "���[�UID"
Public Const LBL_LGI_PASSWORD                   As String = "�p�X���[�h"
Public Const LBL_LGI_STRING                     As String = "�ڑ�������"
Public Const LBL_LGI_CONN_TO                    As String = "�ڑ���"
Public Const LBL_LGI_LOGIN                      As String = "���O�C��"
Public Const LBL_LGI_CANCEL                     As String = "�L�����Z��"


'// ////////////////////
'// �R���{�{�b�N�X (�ϐ�����: CMB_{form code}_{string} )

'// ����
Public Const CMB_COM_CHAR_SET                   As String = "0,S-JIS;1,JIS;2,EUC;3,Unicode(UTF-8)"
Public Const CMB_COM_CR_CODE                    As String = "0,CR(#13) + LF(#10);1,LF(#10);2,CR(#13)"

'// frmCompSheet (CMP)
Public Const CMB_CMP_MARKER                     As String = "0,�������Ȃ�;1,�����𒅐F;2,�Z���𒅐F;3,�g�𒅐F"
Public Const CMB_CMP_METHOD                     As String = "0,�e�L�X�g;1,�l;2,�e�L�X�g�܂��͒l"
Public Const CMB_CMP_OUTPUT                     As String = "0,�ʃu�b�N;1,��r��u�b�N�̖���"

'// frmShowSheetList (SSL)
Public Const CMB_SSL_OUTPUT                     As String = "0,�ʃu�b�N;1,����u�b�N�̐擪;2,����u�b�N�̖���"

'// frmSheetManage (SMG)
Public Const CMB_SMG_TARGET                     As String = "0,���݂̃V�[�g;1,�u�b�N�S��;2,�f�B���N�g���P��"
Public Const CMB_SMG_VIEW                       As String = "0,�w�薳��;1,�W��;2,���y�[�W"
Public Const CMB_SMG_ZOOM                       As String = "0,�w�薳��;1,100;2,75;3,50,4,25"
Public Const CMB_SMG_FILTER                     As String = "0,�w�薳��;1,�t�B���^����;2,�S�ĕ\��;3,1�s�ڂŃt�B���^"

'// frmGetRecord (GRC)
Public Const CMB_GRC_HEADER                     As String = "0,�񖼏̂̂�;1,�񖼏̂ƒ�`;2,�w�b�_����"
Public Const CMB_GRC_GROUP                      As String = "0,�Ȃ�;1,�P��;2,�Q��;3,�R��;4,�S��"

'// frmDataExport (EXP)
'Public Const CMB_EXP_FILE_TYPE                  As String = "0,DML��;1,�Œ蒷�t�@�C��;2,CSV�t�@�C��"
'Public Const CMB_EXP_QUOTE                      As String = "0,"" �_�u���N�H�[�g(#34);1,' �V���O���N�H�[�g(#39);2,�Ȃ�"
'Public Const CMB_EXP_SEPARATOR                  As String = "0,�J���}(#44);1,�Z�~�R����(#59);2,�^�u(#09);3,�X�y�[�X(#32);4,�Ȃ�"
'Public Const CMB_EXP_DATE_FORMAT                As String = "0,yyyy/mm/dd;1,yyyy/mm/dd hh:mm:ss;2,yyyymmdd;3,yyyymmddhhmmss;4,yyyy-mm-dd;5,yyyy-mm-dd hh:mm:ss;6,yyyy-mm-dd-hh.mm.ss"

'// frmXmlManage (XML)
'Public Const CMB_XML_INDENT                     As String = "0,�Ȃ�;1,�^�u(#09);2,�X�y�[�X(#32)�F �Q�o�C�g;3,�X�y�[�X(#32)�F �S�o�C�g;4,�X�y�[�X(#32)�F �W�o�C�g"

'// frmDrawChart (CHT)
Public Const CMB_CHT_POSITION                   As String = "1,�I���Z���̉E;-1,�I���Z���̍�;0,�I���Z����"
Public Const CMB_CHT_DIRECTION                  As String = "0,������;1,�E����"
Public Const CMB_CHT_GRADATION                  As String = "0,�Ȃ�;1,�������̃O���f�[�V����;2,�c�����̃O���f�[�V����(1);3,�c�����̃O���f�[�V����(2)"
Public Const CMB_CHT_INTERVAL                   As String = "0,�Ȃ�;1,#����1;2,#����1;3,#����1;4,#����1;5,#����1;6,#����1;7,#����1;8,#����1;9,#����1"

'// frmOrderShape (ORD)
Public Const CMB_ORD_HEIGHT                     As String = "0,�Z���Ƀt�B�b�g;1,��[����;2,���[����;3,�������Ȃ�"
Public Const CMB_ORD_WIDTH                      As String = "0,�Z���Ƀt�B�b�g;1,���[����;2,�E�[����;3,�������Ȃ�"

'// frmSearch (SRC)
Public Const CMB_SRC_TARGET                     As String = "0,���݂̃V�[�g;1,�u�b�N�S��;2,�t�@�C��"
Public Const CMB_SRC_OUTPUT                     As String = "0,�������Ȃ�;1,�����𒅐F;2,�Z���𒅐F;3,�g�𒅐F"

'// frmFileList (LST)
Public Const CMB_LST_TARGET                     As String = "0,���ׂẴt�@�C��;1,�ȉ��̊g���q�̂�;2,�ȉ��̊g���q�����O"
Public Const CMB_LST_SIZE                       As String = "0,�o�C�g(B);1,�L���o�C�g (KB);2,���K�o�C�g (MB)"
Public Const CMB_LST_DEPTH                      As String = "0,�w��f�B���N�g���̂�;1,1;2,2;3,3;4,4;5,5;6,6;7,7;8,8;9,������"


'// ////////////////////
'// �ꗗ�o�̓w�b�_
Public Const HDR_DISTINCT                       As String = "�l@�J�E���g"   '// �u�J�E���g�v�̕\���񂪉ςȈׁA"@" ��Replace����

'// frmShowSheetList (SSL)
Public Const HDR_SSL                            As String = "�V�[�g�ԍ�;�V�[�g����"

'// frmSearch (SEARCH)
Public Const HDR_SEARCH                         As String = "�t�@�C��;�V�[�g;�Z��;�l;���l"

'// frmFileList (LST)
Public Const HDR_LST                            As String = "�p�X;�t�@�C����;�쐬��;�X�V��;�T�C�Y($);�t�@�C���^�C�v;����;���l"



'// ////////////////////////////////////////////////////////////////////////////
'// 1033 - English
#ElseIf cLANG = 1033 Then



'// ////////////////////
'// �A�v�����ʕϐ� (�ϐ�����: APP_{string} )
Public Const APP_TITLE                          As String = "Excel Extentions"
'Public Const APP_SQL_FILE                       As String = "SQL file (*.sql; *.txt),*.sql;*.txt"
Public Const APP_EXL_FILE                       As String = "Excel file (#),#"
'Public Const APP_XML_FILE                       As String = "XML file (*.xml),*.xml"


'// ////////////////////
'// ���j���[ (�ϐ�����: MENU_{string} )
'Public Const MENU_SHEET_MENU                    As String = "Sheets(&S)"
'Public Const MENU_EXTOOL                        As String = "Extentions(&X)"
'Public Const MENU_SHEET_GROUP                   As String = "Sheet # - @"
'Public Const MENU_COMP_SHEET                    As String = "Difference Check..."
'Public Const MENU_SELECT                        As String = "Select..."
''Public Const MENU_FILE_EXP                      As String = "Data Export..."
'Public Const MENU_SORT                          As String = "Sheet Sort"
'Public Const MENU_SORT_ASC                      As String = "Sort Ascending"
'Public Const MENU_SORT_DESC                     As String = "Sort Descending"
'Public Const MENU_SHEET_LIST                    As String = "Sheet Index..."
'Public Const MENU_SHEET_SETTING                 As String = "Sheet Management..."
Public Const MENU_CHANGE_CHAR                   As String = "Change Case"
Public Const MENU_CAPITAL                       As String = "Uppercase"
Public Const MENU_SMALL                         As String = "Lowercase"
Public Const MENU_PROPER                        As String = "Capital the First Letter in the Word"
Public Const MENU_ZEN                           As String = "Wide Letter"
Public Const MENU_HAN                           As String = "Narrow Letter"
Public Const MENU_TRIM                          As String = "Trim Values"
'Public Const MENU_SELECTION                     As String = "Selection"
'Public Const MENU_SELECTION_UNIT                As String = "Every # Rows"
'Public Const MENU_CLIPBOARD                     As String = "Copy to Clipboard(&C)"
'Public Const MENU_DRAW_LINE_H                   As String = "Draw Header Border"
'Public Const MENU_DRAW_LINE_D                   As String = "Draw Table Border"
'Public Const MENU_RESET                         As String = "Initialize"
'Public Const MENU_VERSION                       As String = "Version Info..."
'Public Const MENU_LINK                          As String = "Hyperlinks"
'Public Const MENU_LINK_ADD                      As String = "Set Hyperlinks"
'Public Const MENU_LINK_REMOVE                   As String = "Remove Hyperlinks"
'Public Const MENU_DRAW_LINE_H_HORIZ             As String = "Top of Table (Horizontal)"
'Public Const MENU_DRAW_LINE_H_VERT              As String = "Left of Table (Vertical)"
''Public Const MENU_XML                           As String = "XML..."
'Public Const MENU_FILE                          As String = "File Index..."
'Public Const MENU_GROUP                         As String = "Group"
'Public Const MENU_GROUP_SET_ROW                 As String = "Grouping (Row)"
'Public Const MENU_GROUP_SET_COL                 As String = "Grouping (Columns)"
'Public Const MENU_GROUP_DISTINCT                As String = "Marge Duplicated Values"
'Public Const MENU_GROUP_VALUE                   As String = "Arrange Duplicated Values"
'Public Const MENU_DRAW_CHART                    As String = "Draw Chart..."
'Public Const MENU_SEARCH                        As String = "Find(&S)..."
'Public Const MENU_RESIZE_SHAPE                  As String = "Fit Shapes to Cell Border..."
'Public Const MENU_SELECTION_BACK_COLR           As String = "Same Background Color"
'Public Const MENU_SELECTION_FONT_COLR           As String = "Same Font Color"


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

'// ���b�Z�[�W�FfrmCompSheet
Public Const MSG_ERROR_NEED_BOOKNAME            As String = "No book identified."
Public Const MSG_NO_FILE                        As String = "Cannot open the file."
Public Const MSG_UNMATCH_SHEET                  As String = "Sheet structure is not the same."
Public Const MSG_NO_DIFF                        As String = "Same data."
Public Const MSG_SHEET_NAME                     As String = "Different sheet name."
Public Const MSG_INS_ROW                        As String = "Inserted"
Public Const MSG_DEL_ROW                        As String = "Removed"

'// ���b�Z�[�W�FfrmSheetManage
Public Const MSG_VAL_10_400                     As String = "Specify the zoom value between 10 and 400."
Public Const MSG_SHEETS_PROTECTED               As String = "Some of the sheets are protected.  Please unprotect them and execute again."
Public Const MSG_COMPLETED_FILES                As String = "The operations on the files below are completed.  Please check the timestamps for confirmation."

'// ���b�Z�[�W�FfrmGetRecord
Public Const MSG_TOO_MANY_ROWS                  As String = "Rows reached Excel limitation.  Further rows omitted."
Public Const MSG_TOO_MANY_COLS                  As String = "Columns reached Excel limitation.  Further columns omitted."
Public Const MSG_QUERY                          As String = "Query to data source"
Public Const MSG_EXTRACT_SHEET                  As String = "Extracting to sheet"
Public Const MSG_PAGE_SETUP                     As String = "Page setup"
Public Const MSG_ROWS_PROCESSED                 As String = " row(s) processed."

'// ���b�Z�[�W�FfrmDataExport
'Public Const MSG_TABLE_NAME                     As String = "Table Name"
'Public Const MSG_COLUMN_NAME                    As String = "Column Name"

'// ���b�Z�[�W�FfrmDrawChart
'Public Const MSG_INVALID_COL_MIN                As String = "Invalid draw position."
'Public Const MSG_INVALID_COL_MAX                As String = "Invalid draw position."

'// ���b�Z�[�W�FfrmLogin
Public Const MSG_LOG_ON_SUCCESS                 As String = "Login successfully."
Public Const MSG_LOG_ON_FAILED                  As String = "Login failed."
Public Const MSG_NEED_FILL_ID                   As String = "User ID required."
Public Const MSG_NEED_FILL_PWD                  As String = "Password required."
Public Const MSG_NEED_FILL_TNS                  As String = "Connection string required."
Public Const MSG_NEED_EXCEL_SAVED               As String = "Current workbook is need to be saved."

'// ���b�Z�[�W�FfrmSearch
Public Const MSG_NO_CONDITION                   As String = "Please specify search condition."
Public Const MSG_WRONG_COND                     As String = "Invalid search condition."

'// ���b�Z�[�W�FfrmFileList
Public Const MSG_MAX_DEPTH                      As String = "Max depth reached."
Public Const MSG_ERR_PRIV                       As String = "Error: Please check your privileges or other settings."
Public Const MSG_EMPTY_DIR                      As String = "Empty"
Public Const MSG_ZERO_BYTE                      As String = "Zero byte file"


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

'// frmCompSheet (CMP)
Public Const LBL_CMP_FORM                       As String = "Difference Check"
Public Const LBL_CMP_MODE_SHEET                 As String = "Sheet"
Public Const LBL_CMP_MODE_BOOK                  As String = "Book"
Public Const LBL_CMP_SHEET1                     As String = "Original"
Public Const LBL_CMP_SHEET2                     As String = "Target"
Public Const LBL_CMP_BOOK1                      As String = "Original"
Public Const LBL_CMP_BOOK2                      As String = "Target"
Public Const LBL_CMP_OPTIONS                    As String = "Options"
Public Const LBL_CMP_RESULT                     As String = "Output to"
Public Const LBL_CMP_MARKER                     As String = "Marker"
Public Const LBL_CMP_METHOD                     As String = "Compare by"
Public Const LBL_CMP_SHOW_COMMENT               As String = "Show comment"

'// frmShowSheetList (SSL)
Public Const LBL_SSL_FORM                       As String = "Create Sheet Index"
Public Const LBL_SSL_TARGET                     As String = "Output to"
Public Const LBL_SSL_OPTIONS                    As String = "Options"
Public Const LBL_SSL_ROWS                       As String = "Rows"
Public Const LBL_SSL_COLS                       As String = "Columns"

'// frmSheetManage (SMG)
Public Const LBL_SMG_FORM                       As String = "Sheet Properties Management"
Public Const LBL_SMG_TARGET                     As String = "Target"
Public Const LBL_SMG_SCROLL                     As String = "Initialize scroll"
Public Const LBL_SMG_FONT_COLOR                 As String = "Initialize font color"
Public Const LBL_SMG_HYPERLINK                  As String = "Remove hyperlinks"
Public Const LBL_SMG_COMMENT                    As String = "Remove comments"
Public Const LBL_SMG_HEAD_FOOT                  As String = "Set header and footer"
Public Const LBL_SMG_MARGIN                     As String = "Set page margin"
Public Const LBL_SMG_PAGEBREAK                  As String = "Remove page breaks"
Public Const LBL_SMG_PRINT_OPT                  As String = "Print settings"
Public Const LBL_SMG_PRINT_NONE                 As String = "None"
Public Const LBL_SMG_PRINT_100                  As String = "100%"
Public Const LBL_SMG_PRINT_HRZ                  As String = "Hrz: one page / Vrt: n pages"
Public Const LBL_SMG_PRINT_1_PAGE               As String = "Print in one page"
Public Const LBL_SMG_VIEW                       As String = "View"
Public Const LBL_SMG_ZOOM                       As String = "Zoom (%)"
Public Const LBL_SMG_AUTOFILTER                 As String = "Auto filter"

'// frmGetRecord (GRC)
Public Const LBL_GRC_FORM                       As String = "Execute SQL Statement"
Public Const LBL_GRC_FILE                       As String = "File"
Public Const LBL_GRC_OPTIONS                    As String = "Options"
Public Const LBL_GRC_DATE_FORMAT                As String = "Date format"
Public Const LBL_GRC_HEADER                     As String = "Header"
Public Const LBL_GRC_GROUPING                   As String = "Grouping"
Public Const LBL_GRC_BORDERS                    As String = "Borders"
Public Const LBL_GRC_BG_COLOR                   As String = "Background color"
Public Const LBL_GRC_SCRIPT                     As String = "SQL script"
Public Const LBL_GRC_LOGIN                      As String = "Login"
Public Const LBL_GRC_FILE_OPEN                  As String = "Open file"
Public Const LBL_GRC_SEARCH                     As String = "Execute"

''// frmDataExport (EXP)
'Public Const LBL_EXP_FORM                       As String = "DML / Data Export"
'Public Const LBL_EXP_FILE_TYPE                  As String = "File type"
'Public Const LBL_EXP_TARGET                     As String = "Target sheets"
'Public Const LBL_EXP_OPTIONS                    As String = "Options"
'Public Const LBL_EXP_DATE_FORMAT                As String = "Date format"
'Public Const LBL_EXP_QUOTE                      As String = "Quotes"
'Public Const LBL_EXP_SEPARATOR                  As String = "Separator"
'Public Const LBL_EXP_CHAR_SET                   As String = "Character set"
'Public Const LBL_EXP_CR_CODE                    As String = "CR code"
'Public Const LBL_EXP_QUOTE_ALL                  As String = "Quote numbers and dates"
'Public Const LBL_EXP_FORMAT_DML                 As String = "Format DML"
'Public Const LBL_EXP_HEADER                     As String = "Add header and footer"
'Public Const LBL_EXP_COL_NAME                   As String = "Column name"
'Public Const LBL_EXP_SEMICOLON                  As String = "Semicolon"
'Public Const LBL_EXP_NUM_POINT                  As String = "Omit decimal point"
'Public Const LBL_EXP_CREATE_SHEET               As String = "New sheet"
'
''// frmXmlManage (XML)
'Public Const LBL_XML_FORM                       As String = "Manage XML File"
'Public Const LBL_XML_INDENT                     As String = "Indent"
'Public Const LBL_XML_PUT_DEF                    As String = "Put XML definition"
'Public Const LBL_XML_LOAD                       As String = "Load XML"
'Public Const LBL_XML_WRITE                      As String = "Write XML"
'
''// frmDrawChart (CHT)
'Public Const LBL_CHT_FORM                       As String = "Draw Chart"
'Public Const LBL_CHT_MAX_VAL                    As String = "Max value"
'Public Const LBL_CHT_INTERVAL                   As String = "Support line intervals"
'Public Const LBL_CHT_POSITION                   As String = "Position"
'Public Const LBL_CHT_DIRECTION                  As String = "Direction"
'Public Const LBL_CHT_GRADATION                  As String = "Gradation"
'Public Const LBL_CHT_LEGEND                     As String = "Show Legend"
'Public Const LBL_CHT_LINE_FRONT                 As String = "Support line in front"

'// frmOrderShape (ORD)
Public Const LBL_ORD_FORM                       As String = "Fit Shapes to Cell Border"
Public Const LBL_ORD_MARGIN                     As String = "Margin"
Public Const LBL_ORD_OPTIONS                    As String = "Options"
Public Const LBL_ORD_HEIGHT                     As String = "Height Setting"
Public Const LBL_ORD_WIDTH                      As String = "Width Setting"

'// frmSearch (SRC)
Public Const LBL_SRC_FORM                       As String = "Find - Advanced Search"
Public Const LBL_SRC_STRING                     As String = "Search string"
Public Const LBL_SRC_TARGET                     As String = "Target"
Public Const LBL_SRC_MARK                       As String = "Marker"
Public Const LBL_SRC_DIR                        As String = "Target Dir"
Public Const LBL_SRC_SUB_DIR                    As String = "Search sub dir"
Public Const LBL_SRC_IGNORE_CASE                As String = "Ignore case"
Public Const LBL_SRC_OBJECT                     As String = "Search in"
Public Const LBL_SRC_CELL_TEXT                  As String = "Cell text"
Public Const LBL_SRC_CELL_FORMULA               As String = "Cell formula"
Public Const LBL_SRC_SHAPE                      As String = "Shape text"
Public Const LBL_SRC_COMMENT                    As String = "Comment"
Public Const LBL_SRC_CELL_NAME                  As String = "Cell name"
Public Const LBL_SRC_SHEET_NAME                 As String = "Sheet name"
Public Const LBL_SRC_HYPERLINK                  As String = "Hyperlink"
Public Const LBL_SRC_HEADER                     As String = "Header / Footer"
Public Const LBL_SRC_GRAPH                      As String = "Graph"

'// frmFileList (LST)
Public Const LBL_LST_FORM                       As String = "File Index"
Public Const LBL_LST_ROOT                       As String = "Dir root"
Public Const LBL_LST_DEPTH                      As String = "Max depth"
Public Const LBL_LST_TARGET                     As String = "Target files"
Public Const LBL_LST_EXT                        As String = "Extentions"
Public Const LBL_LST_SIZE                       As String = "Show size in"
Public Const LBL_LST_REL_PATH                   As String = "Relative path"

'// frmLogin (LGI)
Public Const LBL_LGI_FORM                       As String = "Login"
Public Const LBL_LGI_UID                        As String = "User ID"
Public Const LBL_LGI_PASSWORD                   As String = "Password"
Public Const LBL_LGI_STRING                     As String = "Connection string"
Public Const LBL_LGI_CONN_TO                    As String = "Connect to"
Public Const LBL_LGI_LOGIN                      As String = "Login"
Public Const LBL_LGI_CANCEL                     As String = "Cancel"


'// ////////////////////
'// �R���{�{�b�N�X (�ϐ�����: CMB_{form code}_{string} )

'// ����
Public Const CMB_COM_CHAR_SET                   As String = "0,S-JIS;1,JIS;2,EUC;3,Unicode(UTF-8)"
Public Const CMB_COM_CR_CODE                    As String = "0,CR(#13) + LF(#10);1,LF(#10);2,CR(#13)"

'// frmCompSheet (CMP)
Public Const CMB_CMP_MARKER                     As String = "0,None;1,Font color;2,Background color;3,Border color"
Public Const CMB_CMP_METHOD                     As String = "0,Text;1,Value;2,Text or value"
Public Const CMB_CMP_OUTPUT                     As String = "0,New book;1,The end of the current book"

'// frmShowSheetList (SSL)
Public Const CMB_SSL_OUTPUT                     As String = "0,New Book;1,Top of Current Book;2,End of Current Book"

'// frmSheetManage (SMG)
Public Const CMB_SMG_TARGET                     As String = "0,Current sheet;1,Current book;2,Files in directory"
Public Const CMB_SMG_VIEW                       As String = "0,None;1,Normal;2,Pagebreak"
Public Const CMB_SMG_ZOOM                       As String = "0,None;1,100;2,75;3,50,4,25"
Public Const CMB_SMG_FILTER                     As String = "0,None;1,Disable auto filter;2,Show all;3,Filter with the 1st row"

'// frmGetRecord (GRC)
Public Const CMB_GRC_HEADER                     As String = "0,Column name;1,Column name and type;2,None"
Public Const CMB_GRC_GROUP                      As String = "0,None;1,1 column;2,2 columns;3,3 columns;4,4 columns"

'// frmDataExport (EXP)
'Public Const CMB_EXP_FILE_TYPE                  As String = "0,DML;1,Fixed length;2,CSV"
'Public Const CMB_EXP_QUOTE                      As String = "0,Double quote (#34);1,Single quote (#39);2,None"
'Public Const CMB_EXP_SEPARATOR                  As String = "0,Comma (#44);1,Semicolon (#59);2,Tab (#09);3,Space (#32);4,None"
'Public Const CMB_EXP_DATE_FORMAT                As String = "0,yyyy/mm/dd;1,yyyy/mm/dd hh:mm:ss;2,yyyymmdd;3,yyyymmddhhmmss;4,yyyy-mm-dd;5,yyyy-mm-dd hh:mm:ss;6,yyyy-mm-dd-hh.mm.ss"

'// frmXmlManage (XML)
'Public Const CMB_XML_INDENT                     As String = "0,None;1,Tab (#09);2,Space (#32): 2 bytes;3,Space (#32): 4 bytes;4,Space (#32): 8 bytes"

'// frmDrawChart (CHT)
Public Const CMB_CHT_POSITION                   As String = "1,Right of Selected Area;-1,Left of Selected Area;0,On the Selected Column"
Public Const CMB_CHT_DIRECTION                  As String = "0,Left to Right;1,Right to Left"
Public Const CMB_CHT_GRADATION                  As String = "0,None;1,Horizontal;2,Vertical (1);3,Vertical (2)"
Public Const CMB_CHT_INTERVAL                   As String = "0,None;1,1/#;2,1/#;3,1/#;4,1/#;5,1/#;6,1/#;7,1/#;8,1/#;9,1/#"

'// frmOrderShape (ORD)
Public Const CMB_ORD_HEIGHT                     As String = "0,Fit to cells;1,Top align;2,Bottom align;3,No operation"
Public Const CMB_ORD_WIDTH                      As String = "0,Fit to cells;1,Left align;2,Right align;3,No operation"

'// frmSearch (SRC)
Public Const CMB_SRC_TARGET                     As String = "0,Active sheet;1,Current Book;2,Files"
Public Const CMB_SRC_OUTPUT                     As String = "0,None;1,Text;2,Background color;3,Borders"

'// frmFileList (LST)
Public Const CMB_LST_TARGET                     As String = "0,All files;1,Specified extentions only;2,Exclude specified extentions"
Public Const CMB_LST_SIZE                       As String = "0,Bytes (B);1,Kilobytes (KB);2,Megabytes (MB)"
Public Const CMB_LST_DEPTH                      As String = "0,Current directory only;1,1;2,2;3,3;4,4;5,5;6,6;7,7;8,8;9,Unlimited"


'// ////////////////////
'// �ꗗ�o�̓w�b�_

'// mdlCommon
Public Const HDR_DISTINCT                       As String = "Value@Count"

'// frmShowSheetList (SSL)
Public Const HDR_SSL                            As String = "Sheet Num.;Sheet Name"

'// frmSearch (SEARCH)
Public Const HDR_SEARCH                         As String = "File;Sheet;Cell;Value;Note"

'// frmFileList (LST)
Public Const HDR_LST                            As String = "Location;File Name;Created;Modified;Size($);Type;attributes;Note"



#End If


'// ////////////////////////////////////////////////////////////////////////////
'// END
'// ////////////////////////////////////////////////////////////////////////////
