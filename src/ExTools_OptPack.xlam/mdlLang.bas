Attribute VB_Name = "mdlLang"
'// ////////////////////////////////////////////////////////////////////////////
'// プロジェクト   : 拡張ツール 追加パック
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
Public Const APP_EXL_FILE                       As String = "エクセル形式 ファイル (#),#"


'// ////////////////////
'// メニュー (変数書式: MENU_{string} )


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

'// メッセージ：frmTranslation
Public Const MSG_SERVICE_TRANS_NOT_REACHABLE    As String = "翻訳サイトにアクセスできません。iniファイルの設定とインターネット接続を確認してください。"


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

'// frmTranslation (TRN)
Public Const LBL_TRN_FORM                       As String = "翻訳"
Public Const LBL_TRN_KEY                        As String = "認証キー"
Public Const LBL_TRN_LANG                       As String = "翻訳言語"
Public Const LBL_TRN_OUTPUT                     As String = "出力形式"


'// ////////////////////
'// コンボボックス (変数書式: CMB_{form code}_{string} )

'// 共通
Public Const CMB_COM_CHAR_SET                   As String = "0,S-JIS;1,JIS;2,EUC;3,Unicode(UTF-8)"
Public Const CMB_COM_CR_CODE                    As String = "0,CR(#13) + LF(#10);1,LF(#10);2,CR(#13)"

'// frmTranslation (TRN)
Public Const CMB_TRN_LANGUAGE                   As String = "ja,英語 → 日本語;en,日本語 → 英語"
Public Const CMB_TRN_OUTPUT                     As String = "0,原文の下に追加;1,原文の上に追加;2,原文を削除して上書き;3,右のセルに上書き;4,コメントに追加"


'// ////////////////////
'// 一覧出力ヘッダ



'// ////////////////////////////////////////////////////////////////////////////
'// 1033 - English
#ElseIf cLANG = 1033 Then

'// ////////////////////
'// アプリ共通変数 (変数書式: APP_{string} )
Public Const APP_TITLE                          As String = "Excel Extentions"
Public Const APP_EXL_FILE                       As String = "Excel file (#),#"


'// ////////////////////
'// メニュー (変数書式: MENU_{string} )


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

'// メッセージ：frmTranslation
Public Const MSG_SERVICE_TRANS_NOT_REACHABLE    As String = "Translation service unreachable. Please check your ini file settings and internet connection."


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

'// frmTranslation (TRN)
Public Const LBL_TRN_FORM                       As String = "Translate"
Public Const LBL_TRN_KEY                        As String = "Auth Key"
Public Const LBL_TRN_LANG                       As String = "Target Language"
Public Const LBL_TRN_OUTPUT                     As String = "Output"


'// ////////////////////
'// コンボボックス (変数書式: CMB_{form code}_{string} )

'// 共通
Public Const CMB_COM_CHAR_SET                   As String = "0,S-JIS;1,JIS;2,EUC;3,Unicode(UTF-8)"
Public Const CMB_COM_CR_CODE                    As String = "0,CR(#13) + LF(#10);1,LF(#10);2,CR(#13)"

'// frmTranslation (TRN)
Public Const CMB_TRN_LANGUAGE                   As String = "ja,English → Japanese;en,Japanese → English"
Public Const CMB_TRN_OUTPUT                     As String = "0,Below the original;1,Top of the original;2,Overwrite;3,Right cell;4,Comment"


'// ////////////////////
'// 一覧出力ヘッダ


#End If


'// ////////////////////////////////////////////////////////////////////////////
'// END
'// ////////////////////////////////////////////////////////////////////////////
