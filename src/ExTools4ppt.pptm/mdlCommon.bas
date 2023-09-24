Attribute VB_Name = "mdlCommon"
'// ////////////////////////////////////////////////////////////////////////////
'// プロジェクト   : 拡張ツール
'// タイトル       : 共通関数
'// モジュール     : mdlCommon
'// 説明           : システムの共通関数、起動時の設定などを管理
'// ////////////////////////////////////////////////////////////////////////////
'// Copyright (c) by Koichiro.
'// ////////////////////////////////////////////////////////////////////////////
Option Explicit
Option Base 0

'// ////////////////////////////////////////////////////////////////////////////
'// カスタマイズ可能パラメータ（定数）

Public Const APP_FONT                 As String = "Meiryo UI"                                       '// #001 表示フォント名称
Public Const APP_FONT_SIZE            As Integer = 9                                                '// #002 表示フォントサイズ
Public Const HED_LEFT                 As String = ""                                                '// #003 ヘッダ文字列（左）
Public Const HED_CENTER               As String = ""                                                '// #004 ヘッダ文字列（中央）
Public Const HED_RIGHT                As String = ""                                                '// #005 ヘッダ文字列（右）
Public Const FOT_LEFT                 As String = "&""" & APP_FONT & ",標準""&8&F / &A"             '// #006 フッタ文字列（左）
Public Const FOT_CENTER               As String = "&""" & APP_FONT & ",標準""&8&P / &N"             '// #007 フッタ文字列（中央）
Public Const FOT_RIGHT                As String = "&""" & APP_FONT & ",標準""&8印刷日時: &D &T"     '// #008 フッタ文字列（右）
Public Const MRG_LEFT                 As Double = 0.25                                              '// #009 印刷マージン（左）
Public Const MRG_RIGHT                As Double = 0.25                                              '// #010 印刷マージン（右）
Public Const MRG_TOP                  As Double = 0.75                                              '// #011 印刷マージン（上）
Public Const MRG_BOTTOM               As Double = 0.75                                              '// #012 印刷マージン（下）
Public Const MRG_HEADER               As Double = 0.3                                               '// #013 印刷マージン（ヘッダ）
Public Const MRG_FOOTER               As Double = 0.3                                               '// #014 印刷マージン（フッタ）


'// ////////////////////////////////////////////////////////////////////////////
'// アプリケーション定数

'// バージョン
Public Const APP_VERSION              As String = "3.0.0.77"                                        '// {メジャー}.{機能修正}.{バグ修正}.{開発時管理用}

'// システム定数
Public Const BLANK                    As String = ""                                                '// 空白文字列
Public Const DBQ                      As String = """"                                              '// ダブルクォート
Public Const CHR_ESC                  As Long = 27                                                  '// Escape キーコード
Public Const CLR_ENABLED              As Long = &H80000005                                          '// コントロール背景色 有効
Public Const CLR_DISABLED             As Long = &H8000000F                                          '// コントロール背景色 無効
Public Const TYPE_RANGE               As String = "Range"                                           '// selection タイプ：レンジ
Public Const TYPE_SHAPE               As String = "Shape"                                           '// selection タイプ：シェイプ（varType）
Public Const MENU_PREFIX              As String = "sheet"
Public Const EXCEL_FILE_EXT           As String = "*.xls; *.xlsx"                                   '// エクセル拡張子
Public Const COLOR_ROW                As Integer = 35                                               '// 行色分け色
Public Const COLOR_DIFF_CELL          As Integer = 3                                                '// 色：3=赤
Public Const COLOR_DIFF_ROW_INS       As Integer = 34                                               '// $mod
Public Const COLOR_DIFF_ROW_DEL       As Integer = 15                                               '// $mod
Public Const EXCEL_PASSWORD           As String = ""                                                '// #017 エクセルを開く際のパスワード
Public Const STAT_INTERVAL            As Integer = 100                                              '// ステータスバー更新頻度
Public Const ROW_DIFF_STRIKETHROUGH   As Boolean = True                                             '// $mod
Private Const MENU_NUM                As Integer = 30                                               '// シートをメニューに表示する際のグループ閾値


'// ////////////////////////////////////////////////////////////////////////////
'// パブリック変数

'// 範囲タイプ
Public Type udTargetRange
    minRow  As Long
    minCol  As Integer
    maxRow  As Long
    maxCol  As Integer
    Rows    As Long
    Columns As Integer
End Type

'Public gADO                             As cADO         '// 接続先DB/Excelオブジェクト
Public gLang                            As Long         '// 言語


'// ////////////////////////////////////////////////////////////////////////////
'// メソッド：   ツール初期化
'// 説明：       メニューの構成、アプリオブジェクトの設定を行う。
'// ////////////////////////////////////////////////////////////////////////////
Private Sub psInitExTools()
    '// 言語の設定
    gLang = Application.LanguageSettings.LanguageID(msoLanguageIDInstall)
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// メソッド：   メニュー追加関数
'// 説明：       メニューの追加を行う。親関数（メニュー構成関数）から呼び出される。
'// 引数：       barCtrls:      親バーコントロール
'//              menuCaption:   キャプション
'//              actionCommand: クリック時のイベントプロシージャ
'//              iconNum:       アイコン番号
'//              groupFlag:     グループ線要否
'//              functionID:    パラメータ
'//              menuEnabled:   メニューの有効/無効
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
'// メソッド：   エラーメッセージ表示
'// 説明：       例外処理部で処理できない例外のエラーの内容を、ダイアログ表示する。
'// 引数：       errSource: エラーの発生元のオブジェクトまたはアプリケーションの名前を示す文字列式
'//              e: ＶＢエラーオブジェクト
'//              objAdo： ADOオブジェクト（省略可）
'// ////////////////////////////////////////////////////////////////////////////
Public Sub gsShowErrorMsgDlg(errSource As String, ByVal e As ErrObject)
'Public Sub gsShowErrorMsgDlg(errSource As String, ByVal e As ErrObject, Optional ado As cADO = Nothing)
'    If ado Is Nothing Then
'        '// ADOオブジェクトがからの場合はVBエラーとして扱う
'        Call MsgBox(MSG_ERR & vbLf & vbLf _
'                   & "Error Number: " & e.Number & vbLf _
'                   & "Error Source: " & errSource & vbLf _
'                   & "Error Description: " & e.Description _
'                   , , APP_TITLE)
'        Call e.Clear
'    ElseIf ado.NativeError <> 0 Then
'        '// DBでのエラーの場合
'        Call MsgBox(MSG_ERR & vbLf & vbLf _
'                   & "Error Number: " & ado.NativeError & vbLf _
'                   & "Error Source: Database" & vbLf _
'                   & "Error Description: " & ado.ErrorText _
'                   , , APP_TITLE)
'        ado.InitError
'    ElseIf ado.ErrorCode <> 0 Then
'        '// ADOでのエラーの場合
'        Call MsgBox(MSG_ERR & vbLf & vbLf _
'                   & "Error Number: " & ado.ErrorCode & vbLf _
'                   & "Error Source: ADO" & vbLf _
'                   & "Error Description： " & ado.ErrorText _
'                   , , APP_TITLE)
'        ado.InitError
'    Else
        '// 上記で取り逃した場合はVBエラーとして扱う
        Call MsgBox(MSG_ERR & vbLf & vbLf _
                   & "Error Number: " & e.Number & vbLf _
                   & "Error Source: " & errSource & vbLf _
                   & "Error Description: " & e.Description _
                   , , APP_TITLE)
        Call e.Clear
'    End If
End Sub



'// ////////////////////////////////////////////////////////////////////////////
'// メソッド：   コンボボックス設定
'// 説明：       引数のCSV文字列を基に、コンボボックスの値を設定する。
'// 引数：       targetCombo: 対象コンボボックス
'//              propertyStr: 設定値（{キー},{表示文字列};{キー},{表示文字列}...）
'//              defaultIdx: 初期値
'// ////////////////////////////////////////////////////////////////////////////
'Public Sub gsSetCombo(targetCombo As ComboBox, propertyStr As String, defaultIdx As Integer)
'    Dim lineStr()     As String   '// 設定値の文字列から、各行を格納（;区切り）
'    Dim colStr()      As String   '// 各行の文字列から、列ごとの値を格納（,区切り）
'    Dim idxCnt        As Integer
'
'    lineStr = Split(propertyStr, ";")     '//設定値の文字列を、行毎に分解
'
'    Call targetCombo.Clear
'    For idxCnt = 0 To UBound(lineStr)
'        colStr = Split(lineStr(idxCnt), ",")   '//行の文字列を、カラム毎の文字列に分解
'        Call targetCombo.AddItem(Trim(colStr(0)))
'        targetCombo.List(idxCnt, 1) = Trim(colStr(1))
'    Next
'
'    targetCombo.ListIndex = defaultIdx    '// 初期値を設定
'End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// メソッド：   フォルダ選択ダイアログ表示
'// 説明：       フォルダ選択ダイアログを表示する。
'// 引数：       lngHwnd ウィンドウハンドル
'//              strReturnPath 指定されたフォルダのパス文字列
'// 戻り値：     True:成功  False:失敗(キャンセルを選択した場合含む)
'// ////////////////////////////////////////////////////////////////////////////
'Public Function gfShowSelectFolder(ByVal lngHwnd As Long, ByRef strReturnPath) As Boolean
'    Dim lngRet        As Long
'    Dim lngReturnCode As LongPtr
'    Dim strPath       As String
'    Dim biInfo        As BROWSEINFO
'
'    lngRet = False
'
'    '//文字列領域の確保
'    strPath = String(MAX_PATH + 1, Chr(0))
'
'    ' 構造体の初期化
'    biInfo.hwndOwner = lngHwnd
'    biInfo.lpszTitle = APP_TITLE
'    biInfo.ulFlags = BIF_RETURNONLYFSDIRS
'
'    '// フォルダ選択ダイアログの表示
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
    
'    If ActiveSheet Is Nothing Then                              '// シート（ブック）が開かれているか
'        Call MsgBox(MSG_NO_BOOK, vbOKOnly, APP_TITLE)
'        gfPreCheck = False
'        Exit Function
'    End If
'
'    If protectCont And ActiveSheet.ProtectContents Then         '// アクティブシートが保護されているか
'        Call MsgBox(MSG_SHEET_PROTECTED, vbOKOnly, APP_TITLE)
'        gfPreCheck = False
'        Exit Function
'    End If
'
'    If protectBook And ActiveWorkbook.ProtectStructure Then     '// ブックが保護されているか
'        Call MsgBox(MSG_BOOK_PROTECTED, vbOKOnly, APP_TITLE)
'        gfPreCheck = False
'        Exit Function
'    End If
    
    '// 選択範囲のタイプをチェック
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
    
'    '// 選択範囲カウント
'    If selAreas > 1 Then
'        If Selection.Areas.Count > selAreas Then
'            Call MsgBox(MSG_TOO_MANY_RANGE, vbOKOnly, APP_TITLE)
'            gfPreCheck = False
'            Exit Function
'        End If
'    End If
'
'    '// 選択範囲セルカウント
'    If selCols > 1 Then
'        If Selection.Columns.Count > selCols Then
'            Call MsgBox(MSG_TOO_MANY_COLS_8, vbOKOnly, APP_TITLE)
'            gfPreCheck = False
'            Exit Function
'        End If
'    End If
End Function




'// ////////////////////////////////////////////////////////////////////////////
'// メソッド：   リボンボタンコールバック管理
'// 説明：       リボンからのコールバックをつかさどる
'//              押されたコントロールのIDを基に処理を呼び出す。
'// 引数：       control 対象コントロール
'// ////////////////////////////////////////////////////////////////////////////
Public Sub ribbonCallback(control As IRibbonControl)
    Select Case control.ID
'        '// シート /////
'        Case "SheetComp"                    '// シート比較
'            Call frmCompSheet.Show
'        Case "SheetList"                    '// シート一覧
'            Call frmShowSheetList.Show
'        Case "SheetSetting"                 '// シートの設定
'            Call frmSheetManage.Show
'        Case "SheetSortAsc"                 '// シートの並べ替え
'            Call psSortWorksheet("ASC")
'        Case "SheetSortDesc"                '// シートの並べ替え
'            Call psSortWorksheet("DESC")
'
'        '// データ /////
'        Case "Select"                       '// Select文実行
'            Call frmGetRecord.Show
'
'        '// 値の操作 /////
'        Case "DatePicker"                   '// 日付
'            Call frmDatePicker.Show
'        Case "Today", "Now"                 '// 日付 - 本日日付/現在時刻
'            Call psPutDateTime(control.ID)
'
'        '// 罫線、オブジェクト /////
'        Case "FitObjects"                   '// オブジェクトをセルに合わせる
'            Call frmOrderShape.Show
'        Case "AdjShapeAngle"                '// 円の角度を設定
'            Call frmAdjustArch.Show
'        '// 検索、ファイル /////
'        Case "AdvancedSearch"               '// 拡張検索
'            Call frmSearch.Show
'        Case "FileList"                     '// ファイル一覧
'            Call frmFileList.Show
'
'        '// その他 /////
'        Case "InitTool"                     '// ツール初期化
'            Call psInitExTools
'        Case "Version"                      '// バージョン情報
'            Call frmAbout.Show
    End Select

End Sub


''// ////////////////////////////////////////////////////////////////////////////
''// メソッド：   アプリケーションイベント抑制
''// 説明：       各処理前に再描画や再計算を抑止設定する
''// ////////////////////////////////////////////////////////////////////////////
'Public Sub gsSuppressAppEvents()
'    Application.ScreenUpdating = False                  '// 画面描画停止
'    Application.Cursor = xlWait                         '// ウエイトカーソル
'    Application.EnableEvents = False                    '// イベント抑止
'    If Workbooks.Count > 0 Then
'        Application.Calculation = xlCalculationManual       '// 手動計算
'    End If
'End Sub
'
'
''// ////////////////////////////////////////////////////////////////////////////
''// メソッド：   アプリケーションイベント抑制解除
''// 説明：       各処理後に再描画や再計算を再開する。gsSuppressAppEvents の対
''// ////////////////////////////////////////////////////////////////////////////
'Public Sub gsResumeAppEvents()
'    Application.StatusBar = False                       '// ステータスバーを消す
'    Application.EnableEvents = True
'    Application.Cursor = xlDefault
'    Application.ScreenUpdating = True
'
'    If Workbooks.Count > 0 Then
'        Application.Calculation = xlCalculationAutomatic    '// 自動計算
'    End If
'End Sub





'// ////////////////////////////////////////////////////////////////////////////
'// メソッド：   シェイプ内テキスト取得
'// 説明：       シェイプ内のテキストを取得する。Charactersメソッドをサポートしない場合は例外処理でハンドリング
'//              psExecSearch_Shapeで特定されたシェイプ内のテキストを戻す
'//              V3 からパブリック関数としてfrmSearch → mdlCommon へ移動
'// 引数：       shapeObj: 対象シェイプオブジェクト
'// 戻り値：     シェイプ内のテキスト。シェイプがテキストをサポートしていない場合は一律でブランク
'// ////////////////////////////////////////////////////////////////////////////
'Public Function gfGetShapeText(shapeObj As Shape) As String
'On Error GoTo ErrorHandler
'    If shapeObj.Type = msoTextEffect Then '// ワードアートの場合
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
