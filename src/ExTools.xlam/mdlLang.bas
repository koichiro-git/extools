Attribute VB_Name = "mdlLang"
'// ////////////////////////////////////////////////////////////////////////////
'// プロジェクト   : 拡張ツール
'// タイトル       : ローカライズ設定
'// モジュール     : mdlLang
'// 説明           : 各言語特有の設定
'// ////////////////////////////////////////////////////////////////////////////
'// Copyright (c) by Koichiro.
'// ////////////////////////////////////////////////////////////////////////////
Option Explicit

'// 言語コード Application.LanguageSettings.LanguageID(msoLanguageIDInstall) で得られる値
#Const cLANG = 1041   '// 日本語
'#Const cLANG = 1033  '// English


'// ////////////////////////////////////////////////////////////////////////////
'// 1041 - Japanese
#If cLANG = 1041 Then


'// ////////////////////
'// アプリ共通変数 (変数書式: APP_{string} )
Public Const APP_TITLE                          As String = "拡張ツール"
'Public Const APP_SQL_FILE                       As String = "SQLファイル (*.sql; *.txt),*.sql;*.txt"
Public Const APP_EXL_FILE                       As String = "エクセル形式 ファイル (#),#"
'Public Const APP_XML_FILE                       As String = "XMLファイル (*.xml),*.xml"


'// ////////////////////
'// メニュー (変数書式: MENU_{string} )
'Public Const MENU_SHEET_MENU                    As String = "シート(&S)"
'Public Const MENU_EXTOOL                        As String = "拡張(&X)"
'Public Const MENU_SHEET_GROUP                   As String = "シート # ～ @"
'Public Const MENU_COMP_SHEET                    As String = "シート/ブック比較..."
'Public Const MENU_SELECT                        As String = "SQL文実行..."
''Public Const MENU_FILE_EXP                      As String = "DML/データ出力..."
'Public Const MENU_SORT                          As String = "シートのソート"
'Public Const MENU_SORT_ASC                      As String = "昇順ソート"
'Public Const MENU_SORT_DESC                     As String = "降順ソート"
'Public Const MENU_SHEET_LIST                    As String = "シート一覧を出力..."
'Public Const MENU_SHEET_SETTING                 As String = "シートの設定..."
Public Const MENU_CHANGE_CHAR                   As String = "文字種の変換"
Public Const MENU_CAPITAL                       As String = "大文字"
Public Const MENU_SMALL                         As String = "小文字"
Public Const MENU_PROPER                        As String = "単語の先頭文字を大文字"
Public Const MENU_ZEN                           As String = "全角"
Public Const MENU_HAN                           As String = "半角"
Public Const MENU_TRIM                          As String = "値のトリム"
'Public Const MENU_SELECTION                     As String = "選択範囲の設定"
'Public Const MENU_SELECTION_UNIT                As String = "#行毎"
'Public Const MENU_CLIPBOARD                     As String = "クリップボードへコピー(&C)"
'Public Const MENU_DRAW_LINE_H                   As String = "ヘッダ部の罫線を描画"
'Public Const MENU_DRAW_LINE_D                   As String = "データ部の罫線を描画"
'Public Const MENU_RESET                         As String = "拡張ツールの初期化"
'Public Const MENU_VERSION                       As String = "バージョン情報"
'Public Const MENU_LINK                          As String = "ハイパーリンク"
'Public Const MENU_LINK_ADD                      As String = "ハイパーリンクの設定"
'Public Const MENU_LINK_REMOVE                   As String = "ハイパーリンクの削除"
'Public Const MENU_DRAW_LINE_H_HORIZ             As String = "表の上部（横軸）"
'Public Const MENU_DRAW_LINE_H_VERT              As String = "表の左（縦軸）"
'Public Const MENU_XML                           As String = "XML操作..."
'Public Const MENU_FILE                          As String = "ファイル一覧出力..."
'Public Const MENU_GROUP                         As String = "グループ処理"
'Public Const MENU_GROUP_SET_ROW                 As String = "グループ化（行）"
'Public Const MENU_GROUP_SET_COL                 As String = "グループ化（列）"
'Public Const MENU_GROUP_DISTINCT                As String = "重複データの集約"
'Public Const MENU_GROUP_VALUE                   As String = "重複データを階層風に補正"
'Public Const MENU_DRAW_CHART                    As String = "簡易チャートを描画..."
'Public Const MENU_SEARCH                        As String = "拡張検索(&S)..."
'Public Const MENU_RESIZE_SHAPE                  As String = "シェイプをセルに合わせる..."
'Public Const MENU_SELECTION_BACK_COLR           As String = "同じ背景色"
'Public Const MENU_SELECTION_FONT_COLR           As String = "同じフォント色"


'// ////////////////////
'// メッセージ (変数書式: MSG_{string} )

'// 共通
Public Const MSG_ERR                            As String = "処理中にエラーが発生しました。"
Public Const MSG_NO_BOOK                        As String = "ブックがありません。"
Public Const MSG_FINISHED                       As String = "処理が終了しました。"
Public Const MSG_PROCEED_CREATE_SHEET           As String = "シート作成は、現在のブックにシートを追加します。よろしいですか？"
Public Const MSG_TOO_MANY_RANGE                 As String = "このコマンドは複数の選択範囲に対して実行できません。"
Public Const MSG_NOT_NUMERIC                    As String = "数値を入力してください。"
Public Const MSG_ZERO_NOT_ACCEPTED              As String = "ゼロは入力できません。"
Public Const MSG_NO_DIR                         As String = "検索パスが指定されていません。"
Public Const MSG_DIR_NOT_EXIST                  As String = "指定された検索パスは存在しません。"
Public Const MSG_NO_RESULT                      As String = "検索結果はゼロ件です。"
Public Const MSG_SHEET_PROTECTED                As String = "シートが保護されています。シートの保護を解除して下さい。"
Public Const MSG_BOOK_PROTECTED                 As String = "ブックが保護されています。ブックの保護を解除して下さい。"
Public Const MSG_INVALID_NUM                    As String = "有効な数値を設定してください。"
Public Const MSG_CONFIRM                        As String = "実行します。よろしいですか？"
Public Const MSG_NO_SHEET                       As String = "シートが見つかりません。"
Public Const MSG_INVALID_RANGE                  As String = "有効な値のある範囲を選択してください。"
Public Const MSG_TOO_MANY_COLS_8                As String = "８列以上の選択範囲は処理できません。"
Public Const MSG_NOT_RANGE_SELECT               As String = "セルを選択して下さい。"
Public Const MSG_VLOOKUP_MASTER_2COLS           As String = "VLOOKUPのマスタ表として2列以上を選択して下さい。"
Public Const MSG_VLOOKUP_SET_2COLS              As String = "VLOOKUPの貼り付け先として2列以上を選択して下さい。"
Public Const MSG_VLOOKUP_SEL_DUPLICATED         As String = "VLOOKUPのマスタ表と貼り付け先が重複しています。"
Public Const MSG_VLOOKUP_NO_MASTER              As String = "VLOOKUPのマスタ表が選択されていません。"
Public Const MSG_SEL_DEFAULT_COLOR              As String = "選択されたセルはデフォルト色が指定されています。処理を続けますか？"
Public Const MSG_SHAPE_NOT_SELECTED             As String = "シェイプが選択されていません。"
Public Const MSG_PROCESSING                     As String = "実行中です..."
Public Const MSG_SHAPE_MULTI_SELECT             As String = "2つ以上のシェイプを選択してください。"
Public Const MSG_BARCODE_NOT_AVAILABLE          As String = "バーコードはExcel2016以降で使用可能です。"

'// メッセージ：frmCompSheet
Public Const MSG_ERROR_NEED_BOOKNAME            As String = "ブックを指定してください。"
Public Const MSG_NO_FILE                        As String = "ファイルが開けません"
Public Const MSG_UNMATCH_SHEET                  As String = "シート構成が異なります。"
Public Const MSG_NO_DIFF                        As String = "データは同じです。"
Public Const MSG_SHEET_NAME                     As String = "シート名が異なります"
Public Const MSG_INS_ROW                        As String = "行追加"
Public Const MSG_DEL_ROW                        As String = "行削除"

'// メッセージ：frmSheetManage
Public Const MSG_VAL_10_400                     As String = "ズーム値に有効な数値を入力して下さい。(10～400)"
Public Const MSG_SHEETS_PROTECTED               As String = "保護されているシートが１つ以上あります。シートの保護を解除して下さい。"
Public Const MSG_COMPLETED_FILES                As String = "以下のファイルについての処理は終了しています。念のためファイルの更新日付を確認して下さい。"

'// メッセージ：frmGetRecord
Public Const MSG_TOO_MANY_ROWS                  As String = "出力可能な最大件数に達しました。以降のデータは切り捨てられます。"
Public Const MSG_TOO_MANY_COLS                  As String = "列数が制限値を越えています。制限を越えた列は切り捨てられます。"
Public Const MSG_QUERY                          As String = "問い合わせ中"
Public Const MSG_EXTRACT_SHEET                  As String = "シートへ出力中"
Public Const MSG_PAGE_SETUP                     As String = "書式設定中"
Public Const MSG_ROWS_PROCESSED                 As String = "行が処理されました"

'// メッセージ：frmDataExport
'Public Const MSG_TABLE_NAME                     As String = "表論理名称"
'Public Const MSG_COLUMN_NAME                    As String = "列名称"

'// メッセージ：frmDrawChart
'Public Const MSG_INVALID_COL_MIN                As String = "グラフはA列より左には描画できません。"
'Public Const MSG_INVALID_COL_MAX                As String = "グラフは最大列より右には描画できません。"

'// メッセージ：frmLogin
Public Const MSG_LOG_ON_SUCCESS                 As String = "ログインしました。"
Public Const MSG_LOG_ON_FAILED                  As String = "ログインに失敗しました。"
Public Const MSG_NEED_FILL_ID                   As String = "ユーザＩＤが入力されていません。"
Public Const MSG_NEED_FILL_PWD                  As String = "パスワードが入力されていません。"
Public Const MSG_NEED_FILL_TNS                  As String = "接続文字列が入力されていません。"
Public Const MSG_NEED_EXCEL_SAVED               As String = "現在のブックは保存されていません。ブックを保存してください。"

'// メッセージ：frmSearch
Public Const MSG_NO_CONDITION                   As String = "検索条件を指定して下さい。"
Public Const MSG_WRONG_COND                     As String = "検索条件に指定された文字列は無効です。"

'// メッセージ：frmFileList
Public Const MSG_MAX_DEPTH                      As String = "最大深度に達しました。"
Public Const MSG_ERR_PRIV                       As String = "エラー：アクセス権などに問題がある可能性があります。"
Public Const MSG_EMPTY_DIR                      As String = "空ディレクトリ"
Public Const MSG_ZERO_BYTE                      As String = "ゼロバイト"



'// ////////////////////
'// フォームラベル (変数書式: LBL_{form code}_{string} )

'// 共通
Public Const LBL_COM_EXEC                       As String = "実行"
Public Const LBL_COM_CLOSE                      As String = "閉じる"
Public Const LBL_COM_BROWSE                     As String = "参照..."
Public Const LBL_COM_TARGET                     As String = "出力対象"
Public Const LBL_COM_OPTIONS                    As String = "出力オプション"
Public Const LBL_COM_CHAR_SET                   As String = "文字コード"
Public Const LBL_COM_CR_CODE                    As String = "改行コード"
Public Const LBL_COM_NEW_SHEET                  As String = "シート作成"
Public Const LBL_COM_CHECK_ALL                  As String = "すべて選択"
Public Const LBL_COM_UNCHECK                    As String = "選択解除"
Public Const LBL_COM_HYPERLINK                  As String = "ハイパーリンクの設定"

'// frmCompSheet (CMP)
Public Const LBL_CMP_FORM                       As String = "シート/ブック比較"
Public Const LBL_CMP_MODE_SHEET                 As String = "シート比較"
Public Const LBL_CMP_MODE_BOOK                  As String = "ブック比較"
Public Const LBL_CMP_SHEET1                     As String = "比較元シート"
Public Const LBL_CMP_SHEET2                     As String = "比較先シート"
Public Const LBL_CMP_BOOK1                      As String = "比較元ブック"
Public Const LBL_CMP_BOOK2                      As String = "比較先ブック"
Public Const LBL_CMP_OPTIONS                    As String = "比較オプション"
Public Const LBL_CMP_RESULT                     As String = "出力先"
Public Const LBL_CMP_MARKER                     As String = "マーカー"
Public Const LBL_CMP_METHOD                     As String = "比較方法"
Public Const LBL_CMP_SHOW_COMMENT               As String = "変更箇所のコメントを表示"

'// frmShowSheetList (SSL)
Public Const LBL_SSL_FORM                       As String = "シート一覧出力"
Public Const LBL_SSL_TARGET                     As String = "出力先"
Public Const LBL_SSL_OPTIONS                    As String = "シートの値の出力"
Public Const LBL_SSL_ROWS                       As String = "行数"
Public Const LBL_SSL_COLS                       As String = "列数"

'// frmSheetManage (SMG)
Public Const LBL_SMG_FORM                       As String = "シート操作"
Public Const LBL_SMG_TARGET                     As String = "処理対象"
Public Const LBL_SMG_SCROLL                     As String = "スクロールを初期化"
Public Const LBL_SMG_FONT_COLOR                 As String = "フォント色を初期化"
Public Const LBL_SMG_HYPERLINK                  As String = "ハイパーリンクを削除"
Public Const LBL_SMG_COMMENT                    As String = "コメントを削除"
Public Const LBL_SMG_HEAD_FOOT                  As String = "ヘッダとフッタの表示設定"
Public Const LBL_SMG_MARGIN                     As String = "マージンを設定"
Public Const LBL_SMG_PAGEBREAK                  As String = "改ページと印刷範囲をクリア"
Public Const LBL_SMG_PRINT_OPT                  As String = "印刷の拡大/縮小"
Public Const LBL_SMG_PRINT_NONE                 As String = "設定なし"
Public Const LBL_SMG_PRINT_100                  As String = "100%"
Public Const LBL_SMG_PRINT_HRZ                  As String = "横１枚/縦可変"
Public Const LBL_SMG_PRINT_1_PAGE               As String = "横１枚/縦１枚"
Public Const LBL_SMG_VIEW                       As String = "ビュー"
Public Const LBL_SMG_ZOOM                       As String = "ズーム(%)"
Public Const LBL_SMG_AUTOFILTER                 As String = "オートフィルタ"

'// frmGetRecord (GRC)
Public Const LBL_GRC_FORM                       As String = "SQL文実行"
Public Const LBL_GRC_FILE                       As String = "ファイル"
Public Const LBL_GRC_OPTIONS                    As String = "出力オプション"
Public Const LBL_GRC_DATE_FORMAT                As String = "日付書式"
Public Const LBL_GRC_HEADER                     As String = "ヘッダ出力"
Public Const LBL_GRC_GROUPING                   As String = "グループ化"
Public Const LBL_GRC_BORDERS                    As String = "枠線を表示"
Public Const LBL_GRC_BG_COLOR                   As String = "行を塗り分け"
Public Const LBL_GRC_SCRIPT                     As String = "SQLスクリプト"
Public Const LBL_GRC_LOGIN                      As String = "ログイン"
Public Const LBL_GRC_FILE_OPEN                  As String = "ファイルを開く"
Public Const LBL_GRC_SEARCH                     As String = "実行"

''// frmDataExport (EXP)
'Public Const LBL_EXP_FORM                       As String = "DML/データ出力"
'Public Const LBL_EXP_FILE_TYPE                  As String = "出力形式"
'Public Const LBL_EXP_TARGET                     As String = "出力対象"
'Public Const LBL_EXP_OPTIONS                    As String = "出力オプション"
'Public Const LBL_EXP_DATE_FORMAT                As String = "日付書式"
'Public Const LBL_EXP_QUOTE                      As String = "クォート"
'Public Const LBL_EXP_SEPARATOR                  As String = "区切り文字"
'Public Const LBL_EXP_CHAR_SET                   As String = "文字コード"
'Public Const LBL_EXP_CR_CODE                    As String = "改行コード"
'Public Const LBL_EXP_QUOTE_ALL                  As String = "数値・日付もクォート"
'Public Const LBL_EXP_FORMAT_DML                 As String = "DMLを改行で整形"
'Public Const LBL_EXP_HEADER                     As String = "ヘッダ・フッタを出力"
'Public Const LBL_EXP_COL_NAME                   As String = "項目名を出力"
'Public Const LBL_EXP_SEMICOLON                  As String = "セミコロンを出力しない"
'Public Const LBL_EXP_NUM_POINT                  As String = "数値の小数点を出力しない"
'Public Const LBL_EXP_CREATE_SHEET               As String = "シート作成"
'
''// frmXmlManage (XML)
'Public Const LBL_XML_FORM                       As String = "XML操作"
'Public Const LBL_XML_INDENT                     As String = "インデント"
'Public Const LBL_XML_PUT_DEF                    As String = "XML宣言の出力"
'Public Const LBL_XML_LOAD                       As String = "読込"
'Public Const LBL_XML_WRITE                      As String = "出力"
'
''// frmDrawChart (CHT)
'Public Const LBL_CHT_FORM                       As String = "簡易チャートの描画"
'Public Const LBL_CHT_MAX_VAL                    As String = "最大値"
'Public Const LBL_CHT_INTERVAL                   As String = "補助線間隔"
'Public Const LBL_CHT_POSITION                   As String = "描画位置"
'Public Const LBL_CHT_DIRECTION                  As String = "向き"
'Public Const LBL_CHT_GRADATION                  As String = "グラデーション"
'Public Const LBL_CHT_LEGEND                     As String = "凡例の表示"
'Public Const LBL_CHT_LINE_FRONT                 As String = "補助線を手前に表示"

'// frmOrderShape (ORD)
Public Const LBL_ORD_FORM                       As String = "シェイプの配置"
Public Const LBL_ORD_MARGIN                     As String = "マージン"
Public Const LBL_ORD_OPTIONS                    As String = "詳細設定"
Public Const LBL_ORD_HEIGHT                     As String = "上下幅の設定"
Public Const LBL_ORD_WIDTH                      As String = "左右幅の設定"

'// frmSearch (SRC)
Public Const LBL_SRC_FORM                       As String = "拡張検索"
Public Const LBL_SRC_STRING                     As String = "検索する文字列"
Public Const LBL_SRC_TARGET                     As String = "検索対象"
Public Const LBL_SRC_MARK                       As String = "マーカーの表示"
Public Const LBL_SRC_DIR                        As String = "検索するフォルダ"
Public Const LBL_SRC_SUB_DIR                    As String = "サブフォルダも検索"
Public Const LBL_SRC_IGNORE_CASE                As String = "大文字小文字を区別しない"
Public Const LBL_SRC_OBJECT                     As String = "検索するオブジェクト"
Public Const LBL_SRC_CELL_TEXT                  As String = "セルの文字列を検索"
Public Const LBL_SRC_CELL_FORMULA               As String = "セルの数式を検索"
Public Const LBL_SRC_SHAPE                      As String = "シェイプを検索"
Public Const LBL_SRC_COMMENT                    As String = "コメントを検索"
Public Const LBL_SRC_CELL_NAME                  As String = "セル名称を検索"
Public Const LBL_SRC_SHEET_NAME                 As String = "シート名を検索"
Public Const LBL_SRC_HYPERLINK                  As String = "ハイパーリンクを検索"
Public Const LBL_SRC_HEADER                     As String = "ヘッダ・フッタを検索"
Public Const LBL_SRC_GRAPH                      As String = "グラフを検索"

'// frmFileList (LST)
Public Const LBL_LST_FORM                       As String = "ファイル一覧出力"
Public Const LBL_LST_ROOT                       As String = "出力ルート"
Public Const LBL_LST_DEPTH                      As String = "最大深度"
Public Const LBL_LST_TARGET                     As String = "対象ファイル"
Public Const LBL_LST_EXT                        As String = "拡張子"
Public Const LBL_LST_SIZE                       As String = "サイズ単位"
Public Const LBL_LST_REL_PATH                   As String = "相対パスで表示"

'// frmLogin (LGI)
Public Const LBL_LGI_FORM                       As String = "Login"
Public Const LBL_LGI_UID                        As String = "ユーザID"
Public Const LBL_LGI_PASSWORD                   As String = "パスワード"
Public Const LBL_LGI_STRING                     As String = "接続文字列"
Public Const LBL_LGI_CONN_TO                    As String = "接続先"
Public Const LBL_LGI_LOGIN                      As String = "ログイン"
Public Const LBL_LGI_CANCEL                     As String = "キャンセル"


'// ////////////////////
'// コンボボックス (変数書式: CMB_{form code}_{string} )

'// 共通
Public Const CMB_COM_CHAR_SET                   As String = "0,S-JIS;1,JIS;2,EUC;3,Unicode(UTF-8)"
Public Const CMB_COM_CR_CODE                    As String = "0,CR(#13) + LF(#10);1,LF(#10);2,CR(#13)"

'// frmCompSheet (CMP)
Public Const CMB_CMP_MARKER                     As String = "0,何もしない;1,文字を着色;2,セルを着色;3,枠を着色"
Public Const CMB_CMP_METHOD                     As String = "0,テキスト;1,値;2,テキストまたは値"
Public Const CMB_CMP_OUTPUT                     As String = "0,別ブック;1,比較先ブックの末尾"

'// frmShowSheetList (SSL)
Public Const CMB_SSL_OUTPUT                     As String = "0,別ブック;1,同一ブックの先頭;2,同一ブックの末尾"

'// frmSheetManage (SMG)
Public Const CMB_SMG_TARGET                     As String = "0,現在のシート;1,ブック全体;2,ディレクトリ単位"
Public Const CMB_SMG_VIEW                       As String = "0,指定無し;1,標準;2,改ページ"
Public Const CMB_SMG_ZOOM                       As String = "0,指定無し;1,100;2,75;3,50,4,25"
Public Const CMB_SMG_FILTER                     As String = "0,指定無し;1,フィルタ解除;2,全て表示;3,1行目でフィルタ"

'// frmGetRecord (GRC)
Public Const CMB_GRC_HEADER                     As String = "0,列名称のみ;1,列名称と定義;2,ヘッダ無し"
Public Const CMB_GRC_GROUP                      As String = "0,なし;1,１列;2,２列;3,３列;4,４列"

'// frmDataExport (EXP)
'Public Const CMB_EXP_FILE_TYPE                  As String = "0,DML文;1,固定長ファイル;2,CSVファイル"
'Public Const CMB_EXP_QUOTE                      As String = "0,"" ダブルクォート(#34);1,' シングルクォート(#39);2,なし"
'Public Const CMB_EXP_SEPARATOR                  As String = "0,カンマ(#44);1,セミコロン(#59);2,タブ(#09);3,スペース(#32);4,なし"
'Public Const CMB_EXP_DATE_FORMAT                As String = "0,yyyy/mm/dd;1,yyyy/mm/dd hh:mm:ss;2,yyyymmdd;3,yyyymmddhhmmss;4,yyyy-mm-dd;5,yyyy-mm-dd hh:mm:ss;6,yyyy-mm-dd-hh.mm.ss"

'// frmXmlManage (XML)
'Public Const CMB_XML_INDENT                     As String = "0,なし;1,タブ(#09);2,スペース(#32)： ２バイト;3,スペース(#32)： ４バイト;4,スペース(#32)： ８バイト"

'// frmDrawChart (CHT)
Public Const CMB_CHT_POSITION                   As String = "1,選択セルの右;-1,選択セルの左;0,選択セル上"
Public Const CMB_CHT_DIRECTION                  As String = "0,左から;1,右から"
Public Const CMB_CHT_GRADATION                  As String = "0,なし;1,横方向のグラデーション;2,縦方向のグラデーション(1);3,縦方向のグラデーション(2)"
Public Const CMB_CHT_INTERVAL                   As String = "0,なし;1,#分の1;2,#分の1;3,#分の1;4,#分の1;5,#分の1;6,#分の1;7,#分の1;8,#分の1;9,#分の1"

'// frmOrderShape (ORD)
Public Const CMB_ORD_HEIGHT                     As String = "0,セルにフィット;1,上端揃え;2,下端揃え;3,何もしない"
Public Const CMB_ORD_WIDTH                      As String = "0,セルにフィット;1,左端揃え;2,右端揃え;3,何もしない"

'// frmSearch (SRC)
Public Const CMB_SRC_TARGET                     As String = "0,現在のシート;1,ブック全体;2,ファイル"
Public Const CMB_SRC_OUTPUT                     As String = "0,何もしない;1,文字を着色;2,セルを着色;3,枠を着色"

'// frmFileList (LST)
Public Const CMB_LST_TARGET                     As String = "0,すべてのファイル;1,以下の拡張子のみ;2,以下の拡張子を除外"
Public Const CMB_LST_SIZE                       As String = "0,バイト(B);1,キロバイト (KB);2,メガバイト (MB)"
Public Const CMB_LST_DEPTH                      As String = "0,指定ディレクトリのみ;1,1;2,2;3,3;4,4;5,5;6,6;7,7;8,8;9,無制限"


'// ////////////////////
'// 一覧出力ヘッダ
Public Const HDR_DISTINCT                       As String = "値@カウント"   '// 「カウント」の表示列が可変な為、"@" をReplaceする

'// frmShowSheetList (SSL)
Public Const HDR_SSL                            As String = "シート番号;シート名称"

'// frmSearch (SEARCH)
Public Const HDR_SEARCH                         As String = "ファイル;シート;セル;値;備考"

'// frmFileList (LST)
Public Const HDR_LST                            As String = "パス;ファイル名;作成日;更新日;サイズ($);ファイルタイプ;属性;備考"



'// ////////////////////////////////////////////////////////////////////////////
'// 1033 - English
#ElseIf cLANG = 1033 Then



'// ////////////////////
'// アプリ共通変数 (変数書式: APP_{string} )
Public Const APP_TITLE                          As String = "Excel Extentions"
'Public Const APP_SQL_FILE                       As String = "SQL file (*.sql; *.txt),*.sql;*.txt"
Public Const APP_EXL_FILE                       As String = "Excel file (#),#"
'Public Const APP_XML_FILE                       As String = "XML file (*.xml),*.xml"


'// ////////////////////
'// メニュー (変数書式: MENU_{string} )
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
'// メッセージ (変数書式: MSG_{string} )

'// 共通
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

'// メッセージ：frmCompSheet
Public Const MSG_ERROR_NEED_BOOKNAME            As String = "No book identified."
Public Const MSG_NO_FILE                        As String = "Cannot open the file."
Public Const MSG_UNMATCH_SHEET                  As String = "Sheet structure is not the same."
Public Const MSG_NO_DIFF                        As String = "Same data."
Public Const MSG_SHEET_NAME                     As String = "Different sheet name."
Public Const MSG_INS_ROW                        As String = "Inserted"
Public Const MSG_DEL_ROW                        As String = "Removed"

'// メッセージ：frmSheetManage
Public Const MSG_VAL_10_400                     As String = "Specify the zoom value between 10 and 400."
Public Const MSG_SHEETS_PROTECTED               As String = "Some of the sheets are protected.  Please unprotect them and execute again."
Public Const MSG_COMPLETED_FILES                As String = "The operations on the files below are completed.  Please check the timestamps for confirmation."

'// メッセージ：frmGetRecord
Public Const MSG_TOO_MANY_ROWS                  As String = "Rows reached Excel limitation.  Further rows omitted."
Public Const MSG_TOO_MANY_COLS                  As String = "Columns reached Excel limitation.  Further columns omitted."
Public Const MSG_QUERY                          As String = "Query to data source"
Public Const MSG_EXTRACT_SHEET                  As String = "Extracting to sheet"
Public Const MSG_PAGE_SETUP                     As String = "Page setup"
Public Const MSG_ROWS_PROCESSED                 As String = " row(s) processed."

'// メッセージ：frmDataExport
'Public Const MSG_TABLE_NAME                     As String = "Table Name"
'Public Const MSG_COLUMN_NAME                    As String = "Column Name"

'// メッセージ：frmDrawChart
'Public Const MSG_INVALID_COL_MIN                As String = "Invalid draw position."
'Public Const MSG_INVALID_COL_MAX                As String = "Invalid draw position."

'// メッセージ：frmLogin
Public Const MSG_LOG_ON_SUCCESS                 As String = "Login successfully."
Public Const MSG_LOG_ON_FAILED                  As String = "Login failed."
Public Const MSG_NEED_FILL_ID                   As String = "User ID required."
Public Const MSG_NEED_FILL_PWD                  As String = "Password required."
Public Const MSG_NEED_FILL_TNS                  As String = "Connection string required."
Public Const MSG_NEED_EXCEL_SAVED               As String = "Current workbook is need to be saved."

'// メッセージ：frmSearch
Public Const MSG_NO_CONDITION                   As String = "Please specify search condition."
Public Const MSG_WRONG_COND                     As String = "Invalid search condition."

'// メッセージ：frmFileList
Public Const MSG_MAX_DEPTH                      As String = "Max depth reached."
Public Const MSG_ERR_PRIV                       As String = "Error: Please check your privileges or other settings."
Public Const MSG_EMPTY_DIR                      As String = "Empty"
Public Const MSG_ZERO_BYTE                      As String = "Zero byte file"


'// ////////////////////
'// フォームラベル (変数書式: LBL_{form code}_{string} )

'// 共通
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
'// コンボボックス (変数書式: CMB_{form code}_{string} )

'// 共通
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
'// 一覧出力ヘッダ

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
