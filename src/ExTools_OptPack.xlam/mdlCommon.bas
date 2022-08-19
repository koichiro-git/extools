Attribute VB_Name = "mdlCommon"
'// ////////////////////////////////////////////////////////////////////////////
'// プロジェクト   : 拡張ツール 追加パック
'// タイトル       : 共通関数
'// モジュール     : mdlCommon
'// 説明           : システムの共通関数、起動時の設定などを管理
'// ////////////////////////////////////////////////////////////////////////////
'// Copyright (c) by Koichiro.
'// ////////////////////////////////////////////////////////////////////////////
Option Explicit
Option Base 0

'// ////////////////////////////////////////////////////////////////////////////
'// アプリケーション定数

'// バージョン
Public Const OPTION_PACK_VERSION      As String = "2"                                               '// このモジュール固有のバージョン（管理用通し番号）

'// システム定数
Public Const PROJECT_NAME             As String = "ExToolsOptionalPack"                             '// 本アドイン名称
Public Const BLANK                    As String = ""                                                '// 空白文字列
Public Const CHR_ESC                  As Long = 27                                                  '// Escape キーコード
Public Const CLR_ENABLED              As Long = &H80000005                                          '// コントロール背景色 有効
Public Const CLR_DISABLED             As Long = &H8000000F                                          '// コントロール背景色 無効
Public Const TYPE_RANGE               As String = "Range"                                           '// selection タイプ：レンジ
Public Const TYPE_SHAPE               As String = "Shape"                                           '// selection タイプ：シェイプ（varType）
Public Const MENU_PREFIX              As String = "sheet"


'// ////////////////////////////////////////////////////////////////////////////
'// Windows API 関連の宣言

'// iniファイル読み込み
Declare PtrSafe Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long


'// ////////////////////////////////////////////////////////////////////////////
'// メソッド：   処理実行前チェック（汎用）
'// 説明：       各処理の実行前チェックを行う
'// 引数：
'// 戻り値：     True:成功  False:失敗
'// ////////////////////////////////////////////////////////////////////////////
Public Function gfPreCheck(Optional protectCont As Boolean = False, _
                            Optional protectBook As Boolean = False, _
                            Optional selType As String = BLANK, _
                            Optional selAreas As Integer = 0, _
                            Optional selCols As Integer = 0) As Boolean
  
    gfPreCheck = True
    
    If ActiveSheet Is Nothing Then                              '// シート（ブック）が開かれているか
        Call MsgBox(MSG_NO_BOOK, vbOKOnly, APP_TITLE)
        gfPreCheck = False
        Exit Function
    End If
    
    If protectCont And ActiveSheet.ProtectContents Then         '// アクティブシートが保護されているか
        Call MsgBox(MSG_SHEET_PROTECTED, vbOKOnly, APP_TITLE)
        gfPreCheck = False
        Exit Function
    End If
    
    If protectBook And ActiveWorkbook.ProtectStructure Then     '// ブックが保護されているか
        Call MsgBox(MSG_BOOK_PROTECTED, vbOKOnly, APP_TITLE)
        gfPreCheck = False
        Exit Function
    End If
    
    '// 選択範囲のタイプをチェック
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
    
    '// 選択範囲カウント
    If selAreas > 1 Then
        If Selection.Areas.Count > selAreas Then
            Call MsgBox(MSG_TOO_MANY_RANGE, vbOKOnly, APP_TITLE)
            gfPreCheck = False
            Exit Function
        End If
    End If
    
    '// 選択範囲セルカウント
    If selCols > 1 Then
        If Selection.Columns.Count > selCols Then
            Call MsgBox(MSG_TOO_MANY_COLS_8, vbOKOnly, APP_TITLE)
            gfPreCheck = False
            Exit Function
        End If
    End If
End Function


'// ////////////////////////////////////////////////////////////////////////////
'// メソッド：   アプリケーションイベント抑制
'// 説明：       各処理前に再描画や再計算を抑止設定する
'// ////////////////////////////////////////////////////////////////////////////
Public Sub gsSuppressAppEvents()
    Application.ScreenUpdating = False                  '// 画面描画停止
    Application.Cursor = xlWait                         '// 砂時計カーソル
    Application.EnableEvents = False                    '// イベント抑止
    Application.Calculation = xlCalculationManual       '// 手動計算
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// メソッド：   アプリケーションイベント抑制解除
'// 説明：       各処理後に再描画や再計算を再開する。gsSuppressAppEvents の対
'// ////////////////////////////////////////////////////////////////////////////
Public Sub gsResumeAppEvents()
    Application.StatusBar = False                       '// ステータスバーを消す
    Application.Calculation = xlCalculationAutomatic    '// 自動計算
    Application.EnableEvents = True
    Application.Cursor = xlDefault
    Application.ScreenUpdating = True
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// メソッド：   エラーメッセージ表示（VBA）
'// 説明：       例外処理部で処理できない例外のエラーの内容を、ダイアログ表示する。
'// 引数：       errSource: エラーの発生元のオブジェクトまたはアプリケーションの名前を示す文字列式
'//              e: ＶＢエラーオブジェクト
'//              objAdo： ADOオブジェクト（省略可）
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
'// メソッド：   コンボボックス設定
'// 説明：       引数のCSV文字列を基に、コンボボックスの値を設定する。
'// 引数：       targetCombo: 対象コンボボックス
'//              propertyStr: 設定値（{キー},{表示文字列};{キー},{表示文字列}...）
'//              defaultIdx: 初期値
'// ////////////////////////////////////////////////////////////////////////////
Public Sub gsSetCombo(targetCombo As ComboBox, propertyStr As String, defaultIdx As Integer)
    Dim lineStr()     As String   '// 設定値の文字列から、各行を格納（;区切り）
    Dim colStr()      As String   '// 各行の文字列から、列ごとの値を格納（,区切り）
    Dim idxCnt        As Integer
    
    lineStr = Split(propertyStr, ";")     '//設定値の文字列を、行毎に分解
    
    Call targetCombo.Clear
    For idxCnt = 0 To UBound(lineStr)
        colStr = Split(lineStr(idxCnt), ",")   '//行の文字列を、カラム毎の文字列に分解
        Call targetCombo.AddItem(Trim(colStr(0)))
        targetCombo.List(idxCnt, 1) = Trim(colStr(1))
    Next
    
    targetCombo.ListIndex = defaultIdx    '// 初期値を設定
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// メソッド：   リボンボタンコールバック管理
'// 説明：       リボンからのコールバックをつかさどる
'//              押されたコントロールのIDを基に処理を呼び出す。
'// 引数：       control 対象コントロール
'// ////////////////////////////////////////////////////////////////////////////
Public Sub ribbonCallback2(control As IRibbonControl)
    Select Case control.ID
        Case "FormatPhoneNumbers"                       '// 電話番号補正
            Call gsFormatPhoneNumbers
        Case "Translation"                              '// 翻訳
            Call frmTranslation.Show
    End Select
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// メソッド：   iniファイル設定値取得
'// 説明：       xlamファイルと同名のiniファイルから指定された値を取得する
'// 引数：       section セクション
'//              key     識別キー
'// ////////////////////////////////////////////////////////////////////////////
Public Function gfGetIniFileSetting(section As String, key As String) As String
    Dim sValue      As String   '// 取得バッファ
    Dim lSize       As Long     '// 取得バッファのサイズ
    Dim lRet        As Long     '// 戻り値
    
    '// 取得バッファ初期化
    lSize = 100
    sValue = Space(lSize)
    
    lRet = GetPrivateProfileString(section, key, BLANK, sValue, lSize, Replace(Application.VBE.VBProjects(PROJECT_NAME).Filename, ".xlam", ".ini"))
    gfGetIniFileSetting = Trim(Left(sValue, InStr(sValue, Chr(0)) - 1))
End Function


'// ////////////////////////////////////////////////////////////////////////////
'// END
'// ////////////////////////////////////////////////////////////////////////////
