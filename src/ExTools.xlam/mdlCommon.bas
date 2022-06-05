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

'Public Const COLOR_HEADER             As Integer = 36                                               '// #009 列ヘッダ色
Public Const COLOR_ROW                As Integer = 35                                               '// #018 行色分け色
Public Const COLOR_DIFF_CELL          As Integer = 3                                                '// 色：3=赤
Public Const COLOR_DIFF_ROW_INS       As Integer = 34                                               '// $mod
Public Const COLOR_DIFF_ROW_DEL       As Integer = 15                                               '// $mod
Public Const EXCEL_PASSWORD           As String = ""                                                '// #017 エクセルを開く際のパスワード
Public Const STAT_INTERVAL            As Integer = 100                                              '// ステータスバー更新頻度
Public Const ROW_DIFF_STRIKETHROUGH   As Boolean = True                                             '// $mod
Private Const MAX_COL_LEN             As Integer = 80                                               '// クリップボードにコピーする際の列最大長
Private Const MENU_NUM                As Integer = 30                                               '// シートをメニューに表示する際のグループ閾値


'// ////////////////////////////////////////////////////////////////////////////
'// アプリケーション定数

'// バージョン
Public Const APP_VERSION              As String = "2.1.1.49"                                        '// {メジャー}.{機能修正}.{バグ修正}.{開発時管理用}

'// システム定数
Public Const BLANK                    As String = ""                                                '// 空白文字列
Public Const CHR_ESC                  As Long = 27                                                  '// Escape キーコード
Public Const CLR_ENABLED              As Long = &H80000005                                          '// コントロール背景色 有効
Public Const CLR_DISABLED             As Long = &H8000000F                                          '// コントロール背景色 無効
Public Const TYPE_RANGE               As String = "Range"                                           '// selection タイプ：レンジ
Public Const TYPE_SHAPE               As String = "Shape"                                           '// selection タイプ：シェイプ（varType）
Public Const MENU_PREFIX              As String = "sheet"
Public Const EXCEL_FILE_EXT           As String = "*.xls; *.xlsx"                                   '// エクセル拡張子


'// ////////////////////////////////////////////////////////////////////////////
'// Windows API 関連の宣言

'// 定数
Private Const BIF_RETURNONLYFSDIRS = &H1
Private Const MAX_PATH = 260

'// タイプ
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

'// フォルダ選択
Private Declare Function apiSHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolder" (lpBrowseInfo As BROWSEINFO) As Long
'// パス取得
Private Declare Function apiSHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDList" (ByVal piDL As Long, ByVal strPath As String) As Long
'//キー割り込み
Public Declare Function GetAsyncKeyState Lib "User32.dll" (ByVal vKey As Long) As Long


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
'// プライベート変数
Private pVLookupMaster                  As String               '// VLookUpコピー機能でマスタ表範囲を格納する
Private pVLookupMasterIndex             As String               '// VLookUpコピー機能でマスタ表範囲の表示インデクスを格納する」


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
'// メソッド：   文字種の変換
'// 説明：       選択範囲の値を変換する
'// ////////////////////////////////////////////////////////////////////////////
Private Sub psConvValue(funcFlag As String)
On Error GoTo ErrorHandler
    Dim tCell     As Range    '// 変換対象セル
    Dim statGauge As cStatusGauge
    
    '// 事前チェック（アクティブシート保護、選択タイプ＝セル）
    If Not gfPreCheck(protectCont:=True, selType:=TYPE_RANGE) Then
        Exit Sub
    End If
    
'    Application.ScreenUpdating = False
'    Set statGauge = New cStatusGauge
'    statGauge.MaxVal = Selection.SpecialCells(xlCellTypeConstants, xlNumbers + xlTextValues).Count
    Call gsSuppressAppEvents
    
    If Selection.Count > 1 Then
        For Each tCell In Selection.SpecialCells(xlCellTypeConstants, xlNumbers + xlTextValues)
            Call psConvValue_sub(tCell, funcFlag)
            
            '// キー割込
            If GetAsyncKeyState(27) <> 0 Then
                Application.StatusBar = False
                Exit For
            End If
            
'            Call statGauge.addValue(1)
        Next
    Else
        Call psConvValue_sub(ActiveCell, funcFlag)
    End If
    
'    Set statGauge = Nothing
'    Application.ScreenUpdating = True
    Call gsResumeAppEvents
    Exit Sub
ErrorHandler:
    Call gsResumeAppEvents
    If Err.Number = 1004 Then  '// 範囲選択が正しくない場合
        Call MsgBox(MSG_INVALID_RANGE, vbOKOnly, APP_TITLE)
    Else
        Call gsShowErrorMsgDlg("mdlCommon.psConvValue", Err)
    End If
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// メソッド：   文字種の変換 サブルーチン
'// 説明：       引数の値を変換する
'// ////////////////////////////////////////////////////////////////////////////
Private Sub psConvValue_sub(tCell As Range, funcFlag As String)
    Select Case funcFlag
        Case MENU_CAPITAL
            tCell.Value = UCase(tCell.Value)
        Case MENU_SMALL
            tCell.Value = LCase(tCell.Value)
        Case MENU_PROPER
            tCell.Value = StrConv(tCell.Value, vbProperCase)
        Case MENU_ZEN
            tCell.Value = StrConv(tCell.Value, vbWide)
        Case MENU_HAN
            tCell.Value = StrConv(StrConv(tCell.Value, vbKatakana), vbNarrow)
        Case MENU_TRIM
            tCell.Value = Trim$(tCell.Value)
            If Len(tCell.Value) = 0 Then
                tCell.Value = Empty
            End If
    End Select
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// メソッド：   クリップボードへコピー
'// 説明：       選択範囲を固定長に整形してクリップボードに格納する。
'// 引数：       なし
'// 戻り値：     なし
'// ////////////////////////////////////////////////////////////////////////////
Public Sub psCopyToClipboard()
On Error GoTo ErrorHandler
    Const MAX_LEN   As Integer = 80
    Dim tRange      As udTargetRange
    Dim colLen()    As Integer
    Dim idxRow      As Long
    Dim idxCol      As Integer
    Dim bffText     As String
    Dim rslt        As String
    Dim bffHead     As String
    Dim idxArry     As Integer
    Dim textLen     As Integer
  
    '// 選択範囲が単一であることの確認
    If Selection.Areas.Count > 1 Then
        Call MsgBox(MSG_TOO_MANY_RANGE, vbOKOnly, APP_TITLE)
        Exit Sub
    End If
  
    tRange = gfGetTargetRange(ActiveSheet, Selection)
  
    If (tRange.minRow > tRange.maxRow) Or (tRange.minCol > tRange.maxCol) Then
        Call MsgBox(MSG_INVALID_RANGE, vbOKOnly, APP_TITLE)
        Exit Sub
    End If
  
    '// セルの長さを確認 colLenに格納
    idxArry = 0
    For idxCol = tRange.minCol To tRange.maxCol
        ReDim Preserve colLen(idxArry + 1)
        For idxRow = tRange.minRow To tRange.maxRow
            textLen = LenB(StrConv(WorksheetFunction.Clean(Cells(idxRow, idxCol).Text), vbFromUnicode))
            If textLen > colLen(idxArry) Then
              colLen(idxArry) = textLen
            End If
        Next
        colLen(idxArry) = IIf(colLen(idxArry) = 0, 1, colLen(idxArry))
        colLen(idxArry) = IIf(colLen(idxArry) > MAX_COL_LEN, MAX_COL_LEN, colLen(idxArry))  '// 80バイト以上の長さは切り捨て
        idxArry = idxArry + 1
    Next
  
    For idxRow = tRange.minRow To tRange.maxRow
        For idxCol = 0 To tRange.Columns - 1
            bffText = Trim(WorksheetFunction.Clean(Cells(idxRow, idxCol + tRange.minCol).Text)) '// 改行削除＆Trim。外側のトリムは数値型の場合の符号用空白除去のため必要
            bffText = StrConv(LeftB(StrConv(bffText, vbFromUnicode), 80), vbUnicode) '// 最大文字数以上を足きり
            textLen = LenB(Trim$(StrConv(bffText, vbFromUnicode)))
            If textLen > MAX_LEN Then    '// 80文字以上は切り捨て
                bffText = StrConv(LeftB(StrConv(bffText, vbFromUnicode), colLen(idxCol)), vbUnicode)
            ElseIf IsNumeric(bffText) Or IsDate(bffText) Or pfIsPercentage(bffText) Then    '// 数値、日付は右寄せ
                bffText = Space(colLen(idxCol) - LenB(StrConv(bffText, vbFromUnicode))) & bffText
            Else
                bffText = bffText & Space(colLen(idxCol) - LenB(StrConv(bffText, vbFromUnicode)))
            End If
            rslt = rslt & bffText & Space(1)
        Next
        rslt = Left(rslt, Len(rslt) - 1) & vbCrLf
    Next

    '// 先頭と末尾に罫線を追加
    For idxCol = 0 To tRange.Columns - 1
        bffHead = bffHead & String(colLen(idxCol), "-") & IIf(idxCol = tRange.Columns - 1, vbCrLf, " ")
    Next
    rslt = bffHead & rslt & bffHead
    
    '// クリップボードへコピー ※Win10からDataObjectが動作しなくなるため、回避策SetClipに置き換え
    Call psSetClip(rslt)
'  Set objData = New DataObject
'  Call objData.SetText(rslt)
'  Call objData.PutInClipboard
'  Set objData = Nothing
  
    Exit Sub
ErrorHandler:
    Call gsShowErrorMsgDlg("mdlCommon.psCopyToClipboard", Err)
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// メソッド：   クリップボードへMarkdown形式でコピー
'// 説明：       選択範囲をマークダウン形式でクリップボードに格納する。
'// 引数：       なし
'// 戻り値：     なし
'// ////////////////////////////////////////////////////////////////////////////
Public Sub psCopyToCB_Markdown()
On Error GoTo ErrorHandler
    Dim tRange      As udTargetRange
    Dim colLen()    As Integer
    Dim idxRow      As Long
    Dim idxCol      As Integer
    Dim bffText     As String
    Dim rslt        As String
    Dim bffHead     As String
    Dim idxArry     As Integer
    Dim textLen     As Integer
  
    '// 選択範囲が単一であることの確認
    If Selection.Areas.Count > 1 Then
        Call MsgBox(MSG_TOO_MANY_RANGE, vbOKOnly, APP_TITLE)
        Exit Sub
    End If
  
    tRange = gfGetTargetRange(ActiveSheet, Selection)
  
    If (tRange.minRow > tRange.maxRow) Or (tRange.minCol > tRange.maxCol) Then
        Call MsgBox(MSG_INVALID_RANGE, vbOKOnly, APP_TITLE)
        Exit Sub
    End If
    
    '// ヘッダの出力
    rslt = "|"
    For idxCol = 0 To tRange.Columns - 1
        rslt = rslt & " " & Replace(Cells(tRange.minRow, idxCol + tRange.minCol).Text, vbLf, "<br>") & " |"
    Next
    rslt = rslt & vbCrLf & "|"
    For idxCol = 0 To tRange.Columns - 1
        Select Case Cells(tRange.minRow, idxCol + tRange.minCol).HorizontalAlignment
            Case xlRight
                rslt = rslt & " " & "-: |"
            Case xlCenter
                rslt = rslt & " " & ":-: |"
            Case Else
                rslt = rslt & " " & "- |"
        End Select
    Next
    
    '// データ行の出力
    For idxRow = tRange.minRow + 1 To tRange.maxRow
        rslt = rslt & vbCrLf & "|"
        For idxCol = 0 To tRange.Columns - 1
            rslt = rslt & " " & Replace(Cells(idxRow, idxCol + tRange.minCol).Text, vbLf, "<br>") & " |"
        Next
    Next
    
    '// クリップボードへコピー ※Win10からDataObjectが動作しなくなるため、回避策SetClipに置き換え
    Call psSetClip(rslt)
    
    Exit Sub
ErrorHandler:
    Call gsShowErrorMsgDlg("mdlCommon.psCopyToClipboard_MarkDown", Err)
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// 説明：       Win10 から DataObject.PutInClipboard が効かなくなったため、回避策としてテキストボックスを経由してコピー
'// 引数：       コピー対象文字列
'// ////////////////////////////////////////////////////////////////////////////
Sub psSetClip(bffText As String)
    With CreateObject("Forms.TextBox.1")
        .MultiLine = True
        .Text = bffText
        .SelStart = 0
        .SelLength = .TextLength
        .Copy
    End With
    DoEvents    '// 障害回避のため、一度OSに処理を戻す（再現率低いためこの対処が良いかは確証無い）
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// 説明：       引数の文字列がパーセント形式かを判定する
'// 引数：       コピー対象文字列
Function pfIsPercentage(bffText As String) As Boolean
    If bffText = BLANK Then
        pfIsPercentage = False
    ElseIf Right(bffText, 1) = "%" And IsNumeric(Left(bffText, Len(bffText) - 1)) Then
        pfIsPercentage = True
    Else
        pfIsPercentage = False
    End If
End Function


'// ////////////////////////////////////////////////////////////////////////////
'// メソッド：   選択範囲設定（色による）
'// 説明：       アクティブセルと同じ色のセルを選択範囲に設定する
'// ////////////////////////////////////////////////////////////////////////////
'Private Sub psSetupSelection_color(colorMode As String)
'    Dim targetCell  As Range
'    Dim rgbColor    As Long
'    Dim rslt        As Range
'
'    '// 初期設定
'    If colorMode = "B" Then
'        rgbColor = ActiveCell.Interior.Color
'    Else
'        rgbColor = ActiveCell.Font.Color
'    End If
'
'    '// デフォルト色の場合はキャンセルを促す
'    If (colorMode = "B" And rgbColor = 16777215) _
'      Or (colorMode = "F" And rgbColor = 0) Then
'        If MsgBox(MSG_SEL_DEFAULT_COLOR, vbOKCancel, APP_TITLE) = vbCancel Then
'            Exit Sub
'        End If
'    End If
'
'    Application.ScreenUpdating = False
'    Set rslt = ActiveCell
'    For Each targetCell In ActiveSheet.UsedRange
'        If colorMode = "B" Then
'            If targetCell.Interior.Color = rgbColor Then  '// セル背景色の判定
'                Set rslt = Union(rslt, targetCell)
'            End If
'        Else
'            If targetCell.Font.Color = rgbColor Then      '// フォント色の判定
'                Set rslt = Union(rslt, targetCell)
'            End If
'        End If
'    Next
'
'    Call rslt.Select
'    Application.ScreenUpdating = True
'End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// メソッド：   罫線描画（ヘッダ）
'// 説明：       ヘッダ部の罫線を描画する（横）
'// ////////////////////////////////////////////////////////////////////////////
Public Sub gsDrawLine_Header()
    Dim baseRow As Long     '// 選択領域の開始位置
    Dim baseCol As Integer  '// 選択領域の開始位置
    Dim selRows As Long     '// 選択領域の行数
    Dim selCols As Integer  '// 選択領域の列数
    Dim idxRow  As Long
    Dim idxCol  As Integer
    Dim offRow  As Long
    Dim offCol  As Integer
    Dim childRange As Range
    
    '// 事前チェック（アクティブシート保護、選択タイプ＝セル）
    If Not gfPreCheck(protectCont:=True, selType:=TYPE_RANGE) Then
        Exit Sub
    End If
'    Application.ScreenUpdating = False
    Call gsSuppressAppEvents
    
    For Each childRange In Selection.Areas
        '// 罫線をクリア
        childRange.Borders.LineStyle = xlNone
        childRange.Borders(xlDiagonalDown).LineStyle = xlNone
        childRange.Borders(xlDiagonalUp).LineStyle = xlNone
        
        '// 選択範囲の開始・終了位置取得
        baseRow = childRange.Row
        baseCol = childRange.Column
        selRows = childRange.Rows.Count
        selCols = childRange.Columns.Count
        
        For idxRow = baseRow To baseRow + selRows
            For idxCol = baseCol To baseCol + selCols
                offRow = 0
                offCol = 0
                If (Cells(idxRow, idxCol).Text <> BLANK) Or ((idxRow = baseRow) And (idxCol = baseCol)) Then
                    For offRow = idxRow To baseRow + selRows - 1
                        If (offRow = idxRow) Or Cells(offRow, idxCol).Value = BLANK Then
                            Cells(offRow, idxCol).Borders(xlEdgeLeft).LineStyle = xlContinuous
                        Else
                            Exit For
                        End If
                    Next
                    For offCol = idxCol To baseCol + selCols - 1
                        If (offCol = idxCol) Or Cells(idxRow, offCol).Text = BLANK Then
                            Cells(idxRow, offCol).Borders(xlEdgeTop).LineStyle = xlContinuous
                            If Cells(idxRow, offCol).Borders(xlEdgeRight).LineStyle = xlContinuous Then
                                Exit For
                            End If
                        Else
                            Exit For
                        End If
                    Next
                End If
                '// 最大列に達した場合は終了
                If idxCol = Columns.Count Then
                    Exit For
                End If
            Next
            '// 最大行に達した場合は終了
            If idxRow = Rows.Count Then
                Exit For
            End If
        Next
        
        With childRange.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With
        With childRange.Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With
        'childRange.Interior.ColorIndex = COLOR_HEADER
    Next
'    Application.ScreenUpdating = True
    Call gsResumeAppEvents
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// メソッド：   罫線描画（ヘッダ）：縦
'// 説明：       ヘッダ部の罫線を描画する（縦）
'// 引数：       なし
'// 戻り値：     なし
'// ////////////////////////////////////////////////////////////////////////////
Public Sub gsDrawLine_Header_Vert()
    Dim baseRow As Long     '// 選択領域の開始位置
    Dim baseCol As Integer  '// 選択領域の開始位置
    Dim selRows As Long     '// 選択領域の行数
    Dim selCols As Integer  '// 選択領域の列数
    Dim idxRow  As Long
    Dim idxCol  As Integer
    Dim offRow  As Long
    Dim offCol  As Integer
    Dim childRange As Range
  
    '// 事前チェック（アクティブシート保護、選択タイプ＝セル）
    If Not gfPreCheck(protectCont:=True, selType:=TYPE_RANGE) Then
        Exit Sub
    End If
'    Application.ScreenUpdating = False
    Call gsSuppressAppEvents
    
    For Each childRange In Selection.Areas
        '// 罫線をクリア
        childRange.Borders.LineStyle = xlNone
        childRange.Borders(xlDiagonalDown).LineStyle = xlNone
        childRange.Borders(xlDiagonalUp).LineStyle = xlNone
        
        '// 選択範囲の開始・終了位置取得
        baseRow = childRange.Row
        baseCol = childRange.Column
        selRows = childRange.Rows.Count
        selCols = childRange.Columns.Count
      
        For idxCol = baseCol To baseCol + selCols
            For idxRow = baseRow To baseRow + selRows
                offRow = 0
                offCol = 0
                If (Cells(idxRow, idxCol).Value <> BLANK) Or ((idxRow = baseRow) And (idxCol = baseCol)) Then
                    For offCol = idxCol To baseCol + selCols - 1
                        If (offCol = idxCol) Or Cells(idxRow, offCol).Value = BLANK Then
                            Cells(idxRow, offCol).Borders(xlEdgeTop).LineStyle = xlContinuous
                        Else
                            Exit For
                        End If
                    Next
                    For offRow = idxRow To baseRow + selRows - 1
                        If (offRow = idxRow) Or Cells(offRow, idxCol).Value = BLANK Then
                            Cells(offRow, idxCol).Borders(xlEdgeLeft).LineStyle = xlContinuous
                            If Cells(offRow, idxCol).Borders(xlEdgeBottom).LineStyle = xlContinuous Then
                                Exit For
                            End If
                        Else
                            Exit For
                        End If
                    Next
                End If
                '// 最大行に達した場合は終了
                If idxRow = Rows.Count Then
                    Exit For
                End If
            Next
            '// 最大列に達した場合は終了
            If idxCol = Columns.Count Then
                Exit For
            End If
        Next
    
        With childRange.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With
        With childRange.Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With
    
    'childRange.Interior.ColorIndex = COLOR_HEADER
    Next
    
    Call gsResumeAppEvents
'    Application.ScreenUpdating = True
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// メソッド：   罫線描画（データ）
'// 説明：       データ部の罫線を描画する
'//              選択範囲周辺部をxlThin、内部をxlHairlineで描画する
'// 引数：       なし
'// 戻り値：     なし
'// ////////////////////////////////////////////////////////////////////////////
Public Sub gsDrawLine_Data()
    '// 事前チェック（アクティブシート保護、選択タイプ＝セル）
    If Not gfPreCheck(protectCont:=True, selType:=TYPE_RANGE) Then
        Exit Sub
    End If
    
    Call gsSuppressAppEvents
'    Application.ScreenUpdating = False
    
    '// V2.0より、文字位置は現状のままとする（Selection.VerticalAlignmentは変更しない）よう変更。
    '// 文字位置を上部に設定
'    Selection.VerticalAlignment = xlTop
    
    '// 罫線描画
    Selection.Borders.LineStyle = xlContinuous
    Selection.Borders.Weight = xlThin
    
    If Selection.Columns.Count > 1 Then
        With Selection.Borders(xlInsideVertical)
            .LineStyle = xlContinuous
            .Weight = xlHairline
        End With
    End If
    
    If Selection.Rows.Count > 1 Then
        With Selection.Borders(xlInsideHorizontal)
            .LineStyle = xlContinuous
            .Weight = xlHairline
        End With
    End If
    
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone

'    Application.ScreenUpdating = True
    Call gsResumeAppEvents
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
    Call mdlCommon.gsDrawLine_Data
  
    '// ヘッダの修飾
    If headerLines > 0 Then
        Call wkSheet.Range(wkSheet.Cells(1, 1), wkSheet.Cells(headerLines, wkSheet.UsedRange.Columns.Count)).Select
        Call mdlCommon.gsDrawLine_Header
    
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
'// メソッド：   ハイパーリンクの設定
'// 説明：       選択範囲のハイパーリンクを設定する
'//              標準機能のハイパーリンク設定ではテキスト書式が変わるため、設定前の書式を保持する
'// 引数：       なし
'// 戻り値：     なし
'// ////////////////////////////////////////////////////////////////////////////
Private Sub psSetHyperLink()
    Dim tRange    As udTargetRange
    Dim childRange As Range
    Dim idxRow    As Long
    Dim idxCol    As Integer
    Dim fontName  As String
    Dim fontSize  As String
    Dim fontBold  As Boolean
    Dim fontItlic As Boolean
    Dim fontColor As Double
    
    '// 事前チェック（アクティブシート保護、選択タイプ＝セル）
    If Not gfPreCheck(protectCont:=True, selType:=TYPE_RANGE) Then
        Exit Sub
    End If
    
    Call gsSuppressAppEvents
'    Application.ScreenUpdating = False
    
    For Each childRange In Selection.Areas
        tRange = gfGetTargetRange(ActiveSheet, childRange)
        For idxRow = tRange.minRow To tRange.maxRow
            For idxCol = tRange.minCol To tRange.maxCol
                If Trim(Cells(idxRow, idxCol).Text) <> BLANK Then
                    fontName = Cells(idxRow, idxCol).Font.Name
                    fontSize = Cells(idxRow, idxCol).Font.Size
                    fontBold = Cells(idxRow, idxCol).Font.Bold
                    fontItlic = Cells(idxRow, idxCol).Font.Italic
                    fontColor = Cells(idxRow, idxCol).Font.Color
                    Call Cells(idxRow, idxCol).Hyperlinks.Add(Anchor:=Cells(idxRow, idxCol), Address:=Cells(idxRow, idxCol).Text)
                    Cells(idxRow, idxCol).Font.Name = fontName
                    Cells(idxRow, idxCol).Font.Size = fontSize
                    Cells(idxRow, idxCol).Font.Bold = fontBold
                    Cells(idxRow, idxCol).Font.Italic = fontItlic
                    Cells(idxRow, idxCol).Font.Color = fontColor
                End If
            Next
        Next
    
        '// キー割込
        If GetAsyncKeyState(27) <> 0 Then
            Exit For
        End If
    Next
'    Application.ScreenUpdating = True
    Call gsResumeAppEvents
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// メソッド：   ハイパーリンクの削除
'// 説明：       選択範囲のハイパーリンクを削除する
'// 引数：       なし
'// 戻り値：     なし
'// Excel 2010時点で「HyperLinkのクリア」が標準実装されているが、ツールとしてUIを残すこととした。
'// ////////////////////////////////////////////////////////////////////////////
Private Sub psRemoveHyperLink()
    Dim tRange    As udTargetRange
    Dim idxRow    As Long
    Dim idxCol    As Integer
    Dim fontName  As String
    Dim fontSize  As String
    Dim borderLines(8, 3) As Long
    Dim childRange As Range
    
    '// 事前チェック（アクティブシート保護、選択タイプ＝セル）
    If Not gfPreCheck(protectCont:=True, selType:=TYPE_RANGE) Then
        Exit Sub
    End If
    
    Call Selection.ClearHyperlinks
End Sub


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
'// メソッド：   グループ設定（行）
'// 説明：       グループを自動設定する。
'// 引数：       なし
'// 戻り値：     なし
'// ////////////////////////////////////////////////////////////////////////////
Private Sub psSetGroup_Row()
    Dim idxStart    As Long
    Dim idxEnd      As Long
    Dim idxRow      As Long
    Dim idxCol      As Integer
    Dim tRange      As udTargetRange
    Dim childRange  As Range
    
    '// 事前チェック（アクティブシート保護、選択タイプ＝セル）
    If Not gfPreCheck(protectCont:=True, selType:=TYPE_RANGE) Then
        Exit Sub
    End If
    
'    Application.ScreenUpdating = False
    Call gsSuppressAppEvents
    
  
    '// アウトラインの集計位置を変更
    With ActiveSheet.Outline
        .SummaryRow = xlAbove
    End With
  
        '// グループ設定
        For Each childRange In Selection.Areas
            tRange = gfGetTargetRange(ActiveSheet, childRange)
            
            idxStart = 0
            idxEnd = 0
            idxCol = tRange.minCol
            
            For idxRow = tRange.minRow To tRange.maxRow
                If idxStart = 0 Then
                    idxStart = idxRow + 1
                    idxEnd = idxRow + 1
                ElseIf Trim(Cells(idxRow, idxCol).Text) = BLANK Then
                    idxEnd = idxRow
                ElseIf Trim(Cells(idxRow - 1, idxCol).Text) = BLANK Then
                    Range(Cells(idxStart, 1), Cells(idxEnd, 1)).Rows.Group
                    idxStart = idxRow + 1
                    idxEnd = idxRow + 1
                Else
                    idxStart = idxRow + 1
                    idxEnd = idxRow + 1
                End If
            Next
            If idxStart < idxEnd Then
                Range(Cells(idxStart, 1), Cells(idxEnd, 1)).Rows.Group
            End If
      Next
      
      Call gsResumeAppEvents
'      Application.ScreenUpdating = True
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// メソッド：   グループ設定（列）
'// 説明：       グループを自動設定する。
'// 引数：       なし
'// 戻り値：     なし
'// ////////////////////////////////////////////////////////////////////////////
Private Sub psSetGroup_Col()
    Dim idxStart    As Long
    Dim idxEnd      As Long
    Dim idxRow      As Long
    Dim idxCol      As Integer
    Dim tRange      As udTargetRange
    Dim childRange  As Range
    
    '// 事前チェック（アクティブシート保護、選択タイプ＝セル）
    If Not gfPreCheck(protectCont:=True, selType:=TYPE_RANGE) Then
        Exit Sub
    End If
    
    Call gsSuppressAppEvents
'    Application.ScreenUpdating = False
    
    '// アウトラインの集計位置を変更
    With ActiveSheet.Outline
        .SummaryColumn = xlLeft
    End With
    
    '// グループ設定
    For Each childRange In Selection.Areas
        tRange = gfGetTargetRange(ActiveSheet, childRange)
        
        idxStart = 0
        idxEnd = 0
        idxRow = tRange.minRow
        
        For idxCol = tRange.minCol To tRange.maxCol
            If idxStart = 0 Then
                idxStart = idxCol + 1
                idxEnd = idxCol + 1
            ElseIf Trim(Cells(idxRow, idxCol).Text) = BLANK Then
                idxEnd = idxCol
            ElseIf Trim(Cells(idxRow, idxCol - 1).Text) = BLANK Then
                Range(Cells(1, idxStart), Cells(1, idxEnd)).Columns.Group
                idxStart = idxCol + 1
                idxEnd = idxCol + 1
            Else
                idxStart = idxCol + 1
                idxEnd = idxCol + 1
            End If
        Next
        If idxStart < idxEnd Then
            Range(Cells(1, idxStart), Cells(1, idxEnd)).Columns.Group
        End If
    Next
    
'    Application.ScreenUpdating = True
    Call gsResumeAppEvents
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// メソッド：   値の重複を排除して一覧（カウント）
'// 説明：       重複値を排除する。
'// ////////////////////////////////////////////////////////////////////////////
Private Sub psDistinctVals()
    Dim idxRow      As Long
    Dim idxCol      As Integer
    Dim tRange      As udTargetRange
    
    Dim bff         As Variant
    Dim dict        As Object
    Dim keyString   As String
    Dim keyArray()  As String
    Dim resultSheet As Worksheet
    
    '// セルが選択されていることをチェック
    If TypeName(Selection) <> TYPE_RANGE Then
        Call MsgBox(MSG_NOT_RANGE_SELECT, vbOKOnly, APP_TITLE)
        Exit Sub
    End If
    
    '// チェック
    If Selection.Areas.Count > 1 Then
        Call MsgBox(MSG_TOO_MANY_RANGE, vbOKOnly, APP_TITLE)
        Exit Sub
    End If
    
    Call gsSuppressAppEvents
'    Application.ScreenUpdating = False
    
    tRange = gfGetTargetRange(ActiveSheet, Selection)
    
    bff = Selection.Areas(1)
    Set dict = CreateObject("Scripting.Dictionary")
    
    For idxRow = 1 To tRange.maxRow - tRange.minRow + 1
        '// 行のセルを結合して文字列を一つに
        keyString = BLANK
        For idxCol = 1 To tRange.maxCol - tRange.minCol + 1
            If Not IsError(bff(idxRow, idxCol)) Then
                keyString = keyString & Chr(127) & bff(idxRow, idxCol)
            End If
        Next
        
        If Not dict.Exists(keyString) Then
            Call dict.Add(keyString, "1")
        Else
            dict.Item(keyString) = CStr(CLng(dict.Item(keyString)) + 1)
        End If
    Next
    
    '// 結果出力
    Call Workbooks.Add
    Set resultSheet = ActiveWorkbook.ActiveSheet
    
    '// ヘッダの設定。「カウント」のヘッダ位置を合わせるため、HDR_DISTINCT内の"@"を列数に合わせてReplaceする
    Call gsDrawResultHeader(resultSheet, Replace(HDR_DISTINCT, "@", String(tRange.Columns, ";")), 1)
    
    '// キーの配列をvariantに格納
    bff = dict.Keys
    
    For idxRow = 0 To dict.Count - 1
        keyArray = Split(bff(idxRow), Chr(127))  '// splitは添え字１から開始の仕様？
        For idxCol = 1 To UBound(keyArray)
            resultSheet.Cells(idxRow + 2, idxCol).Value = keyArray(idxCol)
        Next
        
        resultSheet.Cells(idxRow + 2, tRange.maxCol - tRange.minCol + 2).Value = dict.Item(bff(idxRow))
    Next
    
    '//フォント
    Call resultSheet.Cells.Select
    Selection.Font.Name = APP_FONT
    Selection.Font.Size = APP_FONT_SIZE
    
    '// 罫線描画
    Call gsPageSetup_Lines(resultSheet, 1)
    
    '// 閉じるときに保存を求めない
    ActiveWorkbook.Saved = True
    
    Set dict = Nothing
    
'    Application.ScreenUpdating = True
    Call gsResumeAppEvents

End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// メソッド：   値を階層風に補正する
'// 説明：       重複値を階層風に補正する
'// ////////////////////////////////////////////////////////////////////////////
Private Sub psGroupVals()
    Dim idxRow        As Long
    Dim idxCol        As Integer
    Dim tRange        As udTargetRange
    Dim aryIdx        As Integer
    Dim aryLastVal(8) As String
    
    '// セルが選択されていることをチェック
    If TypeName(Selection) <> TYPE_RANGE Then
        Call MsgBox(MSG_NOT_RANGE_SELECT, vbOKOnly, APP_TITLE)
        Exit Sub
    End If
    
    '// チェック
    If Selection.Areas.Count > 1 Then
        Call MsgBox(MSG_TOO_MANY_RANGE, vbOKOnly, APP_TITLE)
        Exit Sub
    ElseIf Selection.Columns.Count > 8 Then
        Call MsgBox(MSG_TOO_MANY_COLS_8, vbOKOnly, APP_TITLE)
        Exit Sub
    End If
    
    Call gsSuppressAppEvents
'    Application.ScreenUpdating = False
    
    tRange = gfGetTargetRange(ActiveSheet, Selection)
    
    For idxRow = tRange.minRow To tRange.maxRow
        For idxCol = tRange.minCol To tRange.maxCol
            If (aryLastVal(idxCol - tRange.minCol) = BLANK) Or (aryLastVal(idxCol - tRange.minCol) <> Cells(idxRow, idxCol).Text) Then
                '// 直前の値が異なる場合 (含 直前の値がブランクの場合)
                '// 配下のレベルの直前の値をクリア
                For aryIdx = tRange.Columns To idxCol Step -1
                    aryLastVal(aryIdx - 1) = BLANK
                Next
                aryLastVal(idxCol - tRange.minCol) = Cells(idxRow, idxCol).Text
            Else
                Cells(idxRow, idxCol).Value = BLANK
            End If
        Next
    Next
    
'    Application.ScreenUpdating = True
    Call gsResumeAppEvents
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// メソッド：   VLookupのマスタ領域としてコピー
'// 説明：       選択領域を表す文字列を内部変数に格納する
'// ////////////////////////////////////////////////////////////////////////////
Private Sub psVLookupCopy()
    If gfPreCheck(selType:=TYPE_RANGE, selAreas:=1) = False Then
        Exit Sub
    End If
    
    '// 1列のみの選択はエラー
    If Selection.Columns.Count = 1 Then
        Call MsgBox(MSG_VLOOKUP_MASTER_2COLS, vbOKOnly, APP_TITLE)
        Exit Sub
    End If
    
    pVLookupMaster = Selection.Worksheet.Name & "!" & Selection.Address(True, True)   '// 列と行を絶対参照
    pVLookupMasterIndex = Selection.Columns.Count
End Sub
            
            
'// ////////////////////////////////////////////////////////////////////////////
'// メソッド：   VLookup関数を張り付け
'// 説明：       VLookupCopyで格納された選択領域を張り付け位置のセルに出力する
'// ////////////////////////////////////////////////////////////////////////////
Private Sub psVLookupPaste()
On Error GoTo ErrorHandler
    Dim searchColIdx    As Long     '// VLookup関数の「検索」にあたるセルの列
    Dim targetColIdx    As Long     '// Vlookpu関数を出力するセルの列
    Dim bffRange        As String   '// 選択範囲のアドレス文字列を保持
    
    '// マスタ列が選択されていない場合はエラー
    If pVLookupMaster = BLANK Then
        Call MsgBox(MSG_VLOOKUP_NO_MASTER, vbOKOnly, APP_TITLE)
        Exit Sub
    End If
    
    '// 1列のみの選択はエラー
    If Selection.Columns.Count = 1 Then
        Call MsgBox(MSG_VLOOKUP_SET_2COLS, vbOKOnly, APP_TITLE)
        Exit Sub
    End If
    
    '// 同一シート内での貼り付けの場合はマスタ表の範囲との重複をチェック
    If Selection.Worksheet.Name = Range(pVLookupMaster).Worksheet.Name Then
        If Not Application.Intersect(Selection, Range(Range(pVLookupMaster).Address)) Is Nothing Then
            Call MsgBox(MSG_VLOOKUP_SEL_DUPLICATED, vbOKOnly, APP_TITLE)
            Exit Sub
        End If
    End If
    
    '// 選択範囲のうち、カレントセルがVLOOKUPの「検索」列に該当
    searchColIdx = ActiveCell.Column
    '// 実際にVLOOKUP関数を埋め込むセルは、選択範囲の最後側。
    targetColIdx = IIf(Selection.Column = ActiveCell.Column, Selection.Column + Selection.Columns.Count - 1, Selection.Column)
    
    '// 最初のセルを出力
    Cells(Selection.Row, targetColIdx).Value = "=VLOOKUP(" & ActiveCell.Address(False, False) & "," & pVLookupMaster & "," & Str(pVLookupMasterIndex) & ",FALSE)"
    '// 下方向に数式のみコピー
    bffRange = Selection.Address(False, False)
    Call Cells(Selection.Row, targetColIdx).Copy
    Call Range(Cells(Selection.Row, targetColIdx), Cells(Selection.Row + Selection.Rows.Count - 1, targetColIdx)).PasteSpecial(Paste:=xlPasteFormulas, Operation:=xlNone, SkipBlanks:=False, Transpose:=False)
    
    '// 後処理
    Application.CutCopyMode = False
    Exit Sub
    
ErrorHandler:
    Call gsShowErrorMsgDlg("mdlCommon.psVLookupPaste", Err)
End Sub


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
    Dim lngReturnCode As Long
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
'// 以下、リボンのコールバック
'// ////////////////////////////////////////////////////////////////////////////


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
        Case "VLookupCopy"                  '// VLookup
            Call psVLookupCopy
        Case "VLookupPaste"
            Call psVLookupPaste
        
        '// 値の操作 /////
        Case "chrUpper"                     '// 大文字
            Call psConvValue(MENU_CAPITAL)
        Case "chrLower"                     '// 小文字
            Call psConvValue(MENU_SMALL)
        Case "chrInitCap"                   '// 先頭大文字
            Call psConvValue(MENU_PROPER)
        Case "chrZen"                       '// 全角
            Call psConvValue(MENU_ZEN)
        Case "chrHan"                       '// 半角
            Call psConvValue(MENU_HAN)
        Case "TrimVal"                      '// トリム
            Call psConvValue(MENU_TRIM)
        Case "AddLink"                      '// リンクの追加
            Call psSetHyperLink
        Case "RemoveLink"                   '// リンクの削除
            Call psRemoveHyperLink
        Case "Copy2Clipboard"               '// 固定長コピー
            Call psCopyToClipboard
        Case "Copy2CBMarkdown"               '// 固定長コピー
            Call psCopyToCB_Markdown
            
        '// 罫線、オブジェクト /////
        Case "groupRow"                     '// グループ化 行
            Call psSetGroup_Row
        Case "groupCol"                     '// グループ化 列
            Call psSetGroup_Col
        Case "removeDup"                    '// 重複のカウント
            Call psDistinctVals
        Case "listDup"                      '// 重複を階層風に補正
            Call psGroupVals
        
        Case "BorderRowHead"                '// 行ヘッダの罫線
            Call gsDrawLine_Header
        Case "BorderColHead"                '// 列ヘッダの罫線
            Call gsDrawLine_Header_Vert
        Case "BorderData"                   '// データ領域の罫線
            Call gsDrawLine_Data
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
'// END
'// ////////////////////////////////////////////////////////////////////////////
