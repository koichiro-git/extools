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
Public Const APP_VERSION              As String = "2.3.2.69"                                        '// {メジャー}.{機能修正}.{バグ修正}.{開発時管理用}

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

Public gADO                             As cADO         '// 接続先DB/Excelオブジェクト
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
Public Sub gsShowErrorMsgDlg(errSource As String, ByVal e As ErrObject, Optional ado As cADO = Nothing)
    If ado Is Nothing Then
        '// ADOオブジェクトがからの場合はVBエラーとして扱う
        Call MsgBox(MSG_ERR & vbLf & vbLf _
                   & "Error Number: " & e.Number & vbLf _
                   & "Error Source: " & errSource & vbLf _
                   & "Error Description: " & e.Description _
                   , , APP_TITLE)
        Call e.Clear
    ElseIf ado.NativeError <> 0 Then
        '// DBでのエラーの場合
        Call MsgBox(MSG_ERR & vbLf & vbLf _
                   & "Error Number: " & ado.NativeError & vbLf _
                   & "Error Source: Database" & vbLf _
                   & "Error Description: " & ado.ErrorText _
                   , , APP_TITLE)
        ado.InitError
    ElseIf ado.ErrorCode <> 0 Then
        '// ADOでのエラーの場合
        Call MsgBox(MSG_ERR & vbLf & vbLf _
                   & "Error Number: " & ado.ErrorCode & vbLf _
                   & "Error Source: ADO" & vbLf _
                   & "Error Description： " & ado.ErrorText _
                   , , APP_TITLE)
        ado.InitError
    Else
        '// 上記で取り逃した場合はVBエラーとして扱う
        Call MsgBox(MSG_ERR & vbLf & vbLf _
                   & "Error Number: " & e.Number & vbLf _
                   & "Error Source: " & errSource & vbLf _
                   & "Error Description: " & e.Description _
                   , , APP_TITLE)
        Call e.Clear
    End If
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// メソッド：   シート並び替え
'// 説明：       シート名で並び替える
'// 引数：       sortMode: 昇順または降順を表す文字列（ASC/DESC）
'// ////////////////////////////////////////////////////////////////////////////
Private Sub psSortWorksheet(sortMode As String)
    Dim i           As Integer
    Dim j           As Integer
    Dim wkSheet     As Worksheet
    Dim isOrderAsc  As Boolean
    
    '// 事前チェック（シート有無）
    If Not gfPreCheck() Then
        Exit Sub
    End If
    
    '// ブックが保護されている場合にはエラーとする
    If ActiveWorkbook.ProtectStructure Then
        Call MsgBox(MSG_BOOK_PROTECTED, vbOKOnly, APP_TITLE)
        Exit Sub
    End If
    
    '// 実行確認
    If MsgBox(MSG_CONFIRM, vbOKCancel, APP_TITLE) = vbCancel Then
        Exit Sub
    End If
    
    Call gsSuppressAppEvents
    
    isOrderAsc = (sortMode = "ASC") '// 昇順/降順の設定
    
    '// ソート
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
    
    '// 後処理
    Call Worksheets(1).Activate
    Call gsResumeAppEvents
    
    Call MsgBox(MSG_FINISHED, vbOKOnly, APP_TITLE)
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// メソッド：    ページ設定(ヘッダ・フッタ)
'// 説明：        ページ設定を行う
'// 引数：        wksheet: ワークシート
'// ////////////////////////////////////////////////////////////////////////////
Public Sub gsPageSetup_Header(wkSheet As Worksheet)
'// $mod プリンタがない場合の明示的なエラーは？
On Error Resume Next
    '// プリンタの設定
    With wkSheet.PageSetup
        '// ヘッダ  ※作成者を表示する場合は右ヘッダのコメントアウト部を使用。
        .LeftHeader = HED_LEFT
        .CenterHeader = HED_CENTER
        .RightHeader = HED_RIGHT
        '// .RightHeader = "&""" & APP_FONT & ",標準""&8作成者:" & Application.UserName & IIf(Application.OrganizationName = BLANK, BLANK, "@" & Application.OrganizationName)
        '// フッタ
        .LeftFooter = FOT_LEFT
        .CenterFooter = FOT_CENTER
        .RightFooter = FOT_RIGHT
    End With
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// メソッド：    ページ設定(マージン)
'// 説明：        マージンの設定を行う
'// 引数：        wksheet: ワークシート
'// ////////////////////////////////////////////////////////////////////////////
Public Sub gsPageSetup_Margin(wkSheet As Worksheet)
On Error Resume Next
    '// プリンタの設定
    With wkSheet.PageSetup
        '// マージン
        .LeftMargin = Application.InchesToPoints(MRG_LEFT)
        .RightMargin = Application.InchesToPoints(MRG_RIGHT)
        .TopMargin = Application.InchesToPoints(MRG_TOP)
        .BottomMargin = Application.InchesToPoints(MRG_BOTTOM)
        .HeaderMargin = Application.InchesToPoints(MRG_HEADER)
        .FooterMargin = Application.InchesToPoints(MRG_FOOTER)
    End With
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// メソッド：    ページ設定(罫線)
'// 説明：        罫線を描画する
'// 引数：        wksheet: ワークシート
'//               headerLines: ヘッダ行数
'// ////////////////////////////////////////////////////////////////////////////
Public Sub gsPageSetup_Lines(wkSheet As Worksheet, headerLines As Integer)
    '// 罫線を描画
    Call wkSheet.UsedRange.Select
    Call gsDrawLine_Data
  
    '// ヘッダの修飾
    If headerLines > 0 Then
        Call wkSheet.Range(wkSheet.Cells(1, 1), wkSheet.Cells(headerLines, wkSheet.UsedRange.Columns.Count)).Select
        Call gsDrawLine_Header
    
        '// ヘッダ下部でウィンドウ枠を固定
        Call wkSheet.Cells(headerLines + 1, 1).Select
        ActiveWindow.FreezePanes = True
    End If
    
    Call wkSheet.Cells(1, 1).Select
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// メソッド：   有効範囲設定
'// 説明：       選択範囲と値の設定されている範囲を比較し、有効範囲を取得する
'// 引数：       wksheet: ワークシート
'//              selRange: 選択範囲
'// 戻り値：     補正後の選択範囲
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
'// メソッド：   列文字列取得
'// 説明：       列の番号を文字表記に変換する
'// 引数：       targetVal: 列番号
'// 戻り値：     列の文字列表記
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
'// メソッド：   セル文字列取得
'// 説明：       text または value プロパティの値を返す
'//              文字列(@)の場合には .Text を戻し、それ以外の場合は $todo
'// 引数：       targetCell: 対象セル
'// 戻り値：     引数のセルの値（text/valueプロパティ）
'// ////////////////////////////////////////////////////////////////////////////
Public Function gfGetTextVal(targetCell As Range) As String
    gfGetTextVal = IIf(targetCell.NumberFormat = "@", targetCell.Value, targetCell.Text)
End Function


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
'// メソッド：   フォルダ選択ダイアログ表示
'// 説明：       フォルダ選択ダイアログを表示する。
'// 引数：       lngHwnd ウィンドウハンドル
'//              strReturnPath 指定されたフォルダのパス文字列
'// 戻り値：     True:成功  False:失敗(キャンセルを選択した場合含む)
'// ////////////////////////////////////////////////////////////////////////////
Public Function gfShowSelectFolder(ByVal lngHwnd As Long, ByRef strReturnPath) As Boolean
    Dim lngRet        As Long
    Dim lngReturnCode As LongPtr
    Dim strPath       As String
    Dim biInfo        As BROWSEINFO
    
    lngRet = False
    
    '//文字列領域の確保
    strPath = String(MAX_PATH + 1, Chr(0))
    
    ' 構造体の初期化
    biInfo.hwndOwner = lngHwnd
    biInfo.lpszTitle = APP_TITLE
    biInfo.ulFlags = BIF_RETURNONLYFSDIRS
    
    '// フォルダ選択ダイアログの表示
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
'// メソッド：   結果シート ヘッダ描画
'// 説明：       引数のヘッダ文字列をシートに出力する
'// 引数：       wkSheet 対象シート
'//              headerStr  出力する文字列
'//              idxRow  出力する行
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
'// メソッド：   シートメニュー get content
'// 説明：       シートのメニュー表示を行う
'// 引数：       control  対象となるリボン上のコントロール
'//              content  戻り値として返す、メニューを表すXML
'// ////////////////////////////////////////////////////////////////////////////
Public Sub sheetMenu_getContent(control As IRibbonControl, ByRef content)
    Dim sheetObj      As Object
    Dim idx           As Integer
    Dim barCtrl_sub   As CommandBarControl
    Dim wkBook        As Workbook
    Dim stMenu        As String
    
    '// $todo:シートが多数ある場合の処理追加
    '// 事前チェック（ブックの有無）
    If Not gfPreCheck() Then
        Exit Sub
    End If
    
    Set wkBook = ActiveWorkbook
    idx = 1
    stMenu = "<menu xmlns=""http://schemas.microsoft.com/office/2006/01/customui"" itemSize=""normal"">"
    
    For Each sheetObj In wkBook.Sheets
        If sheetObj.Type = xlWorksheet Then
            '// IDは接頭辞をつけて通番を設定:MENU_PREFIX + idx
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
'// メソッド：   シートをアクティブにする
'// 説明：       メニューで選択されたメニューキャプションを、アクティブ化の対象にする
'// 引数：       control  押されたシートメニュー。
'// ////////////////////////////////////////////////////////////////////////////
Public Sub sheetMenuOnAction(control As IRibbonControl)
On Error GoTo ErrorHandler
    '// 押されたシートメニューのIDの接頭辞(MENU_PREFIX)を除き、通番をインデックスとして引数に渡す
    Call ActiveWorkbook.Sheets(CInt(Mid(control.ID, Len(MENU_PREFIX) + 1))).Activate
    Exit Sub

ErrorHandler:
    Call MsgBox(MSG_NO_SHEET, vbOKOnly, APP_TITLE)
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// メソッド：   リボンボタンコールバック管理
'// 説明：       リボンからのコールバックをつかさどる
'//              押されたコントロールのIDを基に処理を呼び出す。
'// 引数：       control 対象コントロール
'// ////////////////////////////////////////////////////////////////////////////
Public Sub ribbonCallback(control As IRibbonControl)
    Select Case control.ID
        '// シート /////
        Case "SheetComp"                    '// シート比較
            Call frmCompSheet.Show
        Case "SheetList"                    '// シート一覧
            Call frmShowSheetList.Show
        Case "SheetSetting"                 '// シートの設定
            Call frmSheetManage.Show
        Case "SheetSortAsc"                 '// シートの並べ替え
            Call psSortWorksheet("ASC")
        Case "SheetSortDesc"                '// シートの並べ替え
            Call psSortWorksheet("DESC")
        
        '// データ /////
        Case "Select"                       '// Select文実行
            Call frmGetRecord.Show
        
        '// 値の操作 /////
        Case "DatePicker"                       '// 日付
            Call frmDatePicker.Show
        Case "Today", "Now"                     '// 日付 - 本日日付/現在時刻
            Call psPutDateTime(control.ID)
            
        '// 罫線、オブジェクト /////
        Case "FitObjects"                   '// オブジェクトをセルに合わせる
            Call frmOrderShape.Show
        
        '// 検索、ファイル /////
        Case "AdvancedSearch"               '// 拡張検索
            Call frmSearch.Show
        Case "FileList"                     '// ファイル一覧
            Call frmFileList.Show
        
        '// その他 /////
        Case "InitTool"                     '// ツール初期化
            Call psInitExTools
        Case "Version"                      '// バージョン情報
            Call frmAbout.Show
    End Select

End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// メソッド：   シートをクイックアクセスに表示(Excel2007以降)
'// 説明：       シート一覧をメニューに表示する。
'// ////////////////////////////////////////////////////////////////////////////
Public Sub gsShowSheetOnMenu_2007()
    Dim barCtrl       As CommandBar
    
    '// メニューの初期化
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
'// メソッド：   シートをメニューに表示
'// 説明：       シート一覧をメニューに表示する。
'// 引数：       wkBook: 対象ブック
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
        '// ３０枚以上のシートはグループ化する
        For Each sheetObj In wkBook.Sheets
            If (sheetObj.Index - 1) Mod MENU_NUM = 0 Then
                Set barCtrl_sub = barCtrl.Controls.Add(Type:=msoControlPopup)
                barCtrl_sub.Caption = "シート " & CStr(sheetObj.Index) & " ～ " & CStr(sheetObj.Index + MENU_NUM - 1) & " (&" & IIf(Int(sheetObj.Index / MENU_NUM) < 10, CStr(Int(sheetObj.Index / MENU_NUM)), Chr(55 + Int(sheetObj.Index / MENU_NUM))) & ")"
            End If
            
            If sheetObj.Type = xlWorksheet Then
                Call psPutMenu(barCtrl_sub.Controls, sheetObj.Name & " (&" & pfGetMenuIndex(sheetObj.Index, MENU_NUM) & ")", "psActivateSheet", IIf(sheetObj.ProtectContents, 505, 0), False, sheetObj.Name, (sheetObj.Visible = xlSheetVisible))
            Else '//If (sheetObj.Type = 4) Or (sheetObj.Type = 1) Then
                Call psPutMenu(barCtrl_sub.Controls, sheetObj.Name & " (&" & pfGetMenuIndex(sheetObj.Index, MENU_NUM) & ")", "psActivateSheet", 422, False, sheetObj.Name, (sheetObj.Visible = xlSheetVisible))
            End If
        Next
    Else
        '// ３０枚以下のシートはそのまま表示
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
'// メソッド：   メニューショートカット文字列取得
'// 説明：       シートのメニュー表示にて、ショートカット用文字列を取得する
'// 戻り値：     1～9またはA～Tの文字列
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
'// メソッド：   シートをアクティブにする
'// 説明：       メニューで選択されたメニューキャプションを、アクティブ化の対象にする
'// ////////////////////////////////////////////////////////////////////////////
Public Sub psActivateSheet()
On Error GoTo ErrorHandler
    Call ActiveWorkbook.Sheets(Application.CommandBars.ActionControl.Parameter).Activate
    Exit Sub

ErrorHandler:
    Call MsgBox(MSG_NO_SHEET, vbOKOnly, APP_TITLE)
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// メソッド：   アプリケーションイベント抑制
'// 説明：       各処理前に再描画や再計算を抑止設定する
'// ////////////////////////////////////////////////////////////////////////////
Public Sub gsSuppressAppEvents()
    Application.ScreenUpdating = False                  '// 画面描画停止
    Application.Cursor = xlWait                         '// ウエイトカーソル
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
'// メソッド：   本日日付/現在時刻設定
'// 説明：       アクティブセルに本日日付または現在時刻を設定する
'// ////////////////////////////////////////////////////////////////////////////
Private Sub psPutDateTime(DateTimeMode As String)
    '// 事前チェック（アクティブシート保護、選択タイプ＝セル）
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
    Call gsShowErrorMsgDlg("mdlCommon.psPutDateTime", Err)
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// END
'// ////////////////////////////////////////////////////////////////////////////
