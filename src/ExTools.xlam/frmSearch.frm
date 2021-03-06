VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSearch 
   Caption         =   "拡張検索"
   ClientHeight    =   6375
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6210
   OleObjectBlob   =   "frmSearch.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'// ////////////////////////////////////////////////////////////////////////////
'// プロジェクト   : 拡張ツール
'// タイトル       : 拡張検索フォーム
'// モジュール     : frmSearch
'// 説明           : 正規表現での検索を行う
'// ////////////////////////////////////////////////////////////////////////////
'// Copyright (c) by Koichiro.
'// ////////////////////////////////////////////////////////////////////////////
Option Explicit
Option Base 0


'// ////////////////////////////////////////////////////////////////////////////
'// プライベート変数
'// 検索結果格納タイプ
Private Type udMatched
    FileName    As String
    SheetName   As String
    Row         As Long
    Col         As Integer
    TargetText  As String
    NoteText    As String
    SavedFile   As Boolean
End Type

Private pMatched()  As udMatched    '// 検索結果格納用配列
Private pMatchCnt   As Long         '// 検索結果数


'// //////////////////////////////////////////////////////////////////
'// イベント： フォーム初期化時
Private Sub UserForm_Initialize()
    '// 文字列の検索はデフォルトでON
    ckbSearchText.Value = True
    
    '// コンボボックス設定
    Call gsSetCombo(cmbTarget, CMB_SRC_TARGET, 0)
    Call gsSetCombo(cmbOutput, CMB_SRC_OUTPUT, 0)
    
    '// キャプション設定
    frmSearch.Caption = LBL_SRC_FORM
    cmdDir.Caption = LBL_COM_BROWSE
    ckbSubDir.Caption = LBL_SRC_SUB_DIR
    ckbCaseSensitive.Caption = LBL_SRC_IGNORE_CASE
    fraOptions.Caption = LBL_SRC_OBJECT
    ckbSearchText.Caption = LBL_SRC_CELL_TEXT
    ckbSearchFormula.Caption = LBL_SRC_CELL_FORMULA
    ckbSearchShape.Caption = LBL_SRC_SHAPE
    ckbSearchComment.Caption = LBL_SRC_COMMENT
    ckbSearchName.Caption = LBL_SRC_CELL_NAME
    ckbSearchSheetName.Caption = LBL_SRC_SHEET_NAME
    ckbSearchLink.Caption = LBL_SRC_HYPERLINK
    ckbSearchHeader.Caption = LBL_SRC_HEADER
    ckbSearchGraph.Caption = LBL_SRC_GRAPH
    lblString.Caption = LBL_SRC_STRING
    lblTarget.Caption = LBL_SRC_TARGET
    lblMarker.Caption = LBL_SRC_MARK
    lblDir.Caption = LBL_SRC_DIR
    cmdSelectAll.Caption = LBL_COM_CHECK_ALL
    cmdClear.Caption = LBL_COM_UNCHECK
    cmdExecute.Caption = LBL_COM_EXEC
    cmdClose.Caption = LBL_COM_CLOSE
End Sub


'// //////////////////////////////////////////////////////////////////
'// イベント： 閉じるボタン クリック時
Private Sub cmdClose_Click()
    Call Me.Hide
End Sub


'// //////////////////////////////////////////////////////////////////
'// イベント： 参照ボタン クリック時
Private Sub cmdDir_Click()
    Dim FilePath  As String
    
    If Not gfShowSelectFolder(0, FilePath) Then
        Exit Sub
    Else
        txtDirectory.Text = FilePath
    End If
End Sub


'// //////////////////////////////////////////////////////////////////
'// イベント： 検索対象コンボ 変更時
Private Sub cmbTarget_Change()
    Select Case cmbTarget.Value
        Case 0  '// 現在のシート
            cmdDir.Enabled = False
            ckbSubDir.Enabled = False
            txtDirectory.Enabled = False
            txtDirectory.BackColor = CLR_DISABLED
            ckbSearchSheetName.Enabled = False
            cmbOutput.Enabled = True
        Case 1  '// ブック全体
            cmdDir.Enabled = False
            ckbSubDir.Enabled = False
            txtDirectory.Enabled = False
            txtDirectory.BackColor = CLR_DISABLED
            ckbSearchSheetName.Enabled = True
            cmbOutput.Enabled = True
        Case 2  '// ディレクトリ単位
            cmdDir.Enabled = True
            ckbSubDir.Enabled = True
            txtDirectory.Enabled = True
            txtDirectory.BackColor = CLR_ENABLED
            ckbSearchSheetName.Enabled = True
            cmbOutput.Enabled = False
    End Select
End Sub


'// //////////////////////////////////////////////////////////////////
'// イベント： 実行ボタン クリック時
Private Sub cmdExecute_Click()
    Dim wkSheet   As Worksheet
    Dim fs        As Object
  
    '// nullチェック
    If Trim(txtSearch.Value) = BLANK Then
      Call MsgBox(MSG_NO_CONDITION, vbOKOnly, APP_TITLE)
      Call txtSearch.SetFocus
      Exit Sub
    End If
  
    Call gsSuppressAppEvents
    
    '// 検索結果クリア
    pMatchCnt = 0
    Erase pMatched
    
    '// 検索実行（psExecSearch呼び出し）
    Select Case cmbTarget.Value
        Case 0  '// 現在のシート
            Call psExecSearch(ActiveSheet, txtSearch.Text, ckbCaseSensitive.Value)
        Case 1  '// ブック全体
            For Each wkSheet In ActiveWorkbook.Sheets
                Call psExecSearch(wkSheet, txtSearch.Text, ckbCaseSensitive.Value)
            Next
        Case 2  '// ディレクトリ単位
            If Trim(txtDirectory.Text) <> BLANK Then
                Set fs = CreateObject("Scripting.FileSystemObject")
                
                '// 検索パス確認
                If fs.FolderExists(txtDirectory.Text) Then
                    Call psGetExcelFiles(fs, txtDirectory.Text, txtSearch.Text, ckbCaseSensitive.Value, ckbSubDir.Value)
                Else
                    Call MsgBox(MSG_DIR_NOT_EXIST, vbOKOnly, APP_TITLE)
                    Exit Sub
                End If
                Set fs = Nothing
            Else
                Call MsgBox(MSG_NO_DIR, vbOKOnly, APP_TITLE)
                Call txtDirectory.SetFocus
                Application.ScreenUpdating = True
                Exit Sub
            End If
    End Select
    
    If pMatchCnt > 0 Then   '// 検索結果が1件以上あればシートに出力し、処理完了
        Call psShowResult
        Call MsgBox(MSG_FINISHED, vbOKOnly, APP_TITLE)
        Call Me.Hide
    Else
        Call MsgBox(MSG_NO_RESULT, vbOKOnly, APP_TITLE)
    End If
  
    Application.StatusBar = False
    Application.ScreenUpdating = True
End Sub


'// //////////////////////////////////////////////////////////////////
'// イベント： 全てを選択ボタン クリック時
Private Sub cmdSelectAll_Click()
    Call psSetCheckBoxes(True)
End Sub


'// //////////////////////////////////////////////////////////////////
'// イベント： 選択解除ボタン クリック時
Private Sub cmdClear_Click()
    Call psSetCheckBoxes(False)
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// メソッド：   検索対象チェックボックス設定
'// 説明：       検索対象チェックボックスの値を引数の真偽値に一括設定する。
'// 引数：       newValue: チェックボックスの設定値
'// 戻り値：     なし
'// ////////////////////////////////////////////////////////////////////////////
Private Sub psSetCheckBoxes(newValue As Boolean)
    ckbSearchText.Value = newValue
    ckbSearchFormula.Value = newValue
    ckbSearchShape.Value = newValue
    ckbSearchComment.Value = newValue
    ckbSearchName.Value = newValue
    ckbSearchSheetName.Value = newValue
    ckbSearchLink.Value = newValue
    ckbSearchHeader.Value = newValue
    ckbSearchGraph.Value = newValue
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// メソッド：   ディレクトリ内ブック検索
'// 説明：       指定されたディレクトリ内のブックを検索する
'// 引数：       fs: ファイルシステムオブジェクト
'//              dirName: 検索対象ディレクトリ
'//              patternStr: 検索文字列
'//              caseSensitive: 大文字小文字の区別フラグ
'//              searchSubDir: サブディレクトリ検索フラグ
'// 戻り値：     なし
'// ////////////////////////////////////////////////////////////////////////////
Private Sub psGetExcelFiles(fs As Object, dirName As String, patternStr As String, caseSensitive As Boolean, searchSubDir As Boolean)
    Dim parentDir   As Object
    Dim children    As Object
    Dim wkBook      As Workbook
    Dim wkSheet     As Worksheet
    Dim isDuplName  As Boolean    '// 対象となるブックが開かれている場合True
    
    Set parentDir = fs.GetFolder(dirName)
    
    '// ファイルの検索
    For Each children In parentDir.files
        With children
            If LCase(Right(.Name, 3)) = "xls" Then
                '// 検索
                '// ブックが既に開かれているかを確認
                isDuplName = False
                For Each wkBook In Workbooks
                    If wkBook.Name = children.Name Then
                        isDuplName = True
                        Exit For
                    End If
                Next
                
                If Not isDuplName Then  '// ブックが開かれている場合は検索対象外
                    Set wkBook = Workbooks.Open(children.Path, ReadOnly:=True, password:=EXCEL_PASSWORD)
                    For Each wkSheet In wkBook.Worksheets
                        Call psExecSearch(wkSheet, patternStr, caseSensitive)
                    Next
                    Call wkBook.Close(SaveChanges:=False)
                    Set wkBook = Nothing
                End If
            End If
        End With
    Next
    
    '// サブフォルダがある場合、検索
    If searchSubDir Then
        For Each children In parentDir.SubFolders
          '// 子ディレクトリの再帰呼び出し
          Call psGetExcelFiles(fs, children.Path, patternStr, caseSensitive, True)
        Next
    End If
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// メソッド：   ディレクトリ内ブック検索
'// 説明：       指定されたディレクトリ内のブックを検索する
'// 引数：       wkSheet: 検索対象シート
'//              patternStr: 検索文字列
'//              caseSensitive: 大文字小文字の区別フラグ
'// 戻り値：     なし
'// ////////////////////////////////////////////////////////////////////////////
Private Sub psExecSearch(wkSheet As Worksheet, patternStr As String, caseSensitive As Boolean)
    Dim regExp        As Object         '// 正規表現オブジェクト
    Dim targetCell    As Range
    Dim hLink         As Hyperlink
    Dim rangeName     As Name
    Dim shapeObj      As Shape
    Dim commentObj    As Comment
    Dim chartObj      As Chart
    Dim seriesObj     As Series
    Dim bffText       As String
    Dim idxChart      As Long
    Dim idxCellSrch   As Long           '// 検索セル数カウンタ
    Dim numCellCnt    As Long           '// 検索対象セル数
  
    numCellCnt = numCellCnt + IIf(ckbSearchText.Value, wkSheet.UsedRange.Count, 0)
    If pfGetCellCount(wkSheet.UsedRange, xlCellTypeFormulas) > -1 Then
        numCellCnt = numCellCnt + IIf(ckbSearchFormula.Value, wkSheet.UsedRange.SpecialCells(xlCellTypeFormulas).Count, 0)
    End If
  
    '// 正規表現オブジェクトの作成
    Set regExp = CreateObject("VBScript.RegExp")
    regExp.Pattern = patternStr
    regExp.IgnoreCase = caseSensitive
  
    '// 検索試行（正規表現の記載チェック）
    If Not pfCheckRegExp(regExp) Then
        Call MsgBox(MSG_WRONG_COND, vbOKOnly, APP_TITLE)
        Set regExp = Nothing
        Exit Sub
    End If
  
    '// セル文字列を検索 //////////
    If ckbSearchText.Value Then
        For Each targetCell In wkSheet.UsedRange
            If regExp.test(targetCell.Text) Then
                Call psSetMatchedRec(wkSheet, targetCell.Row, targetCell.Column, targetCell.Text, BLANK)
                
                '// セル着色など
                Select Case cmbOutput.Value
                    Case 0  '// 何もしない
                    Case 1  '// 文字を着色
                      targetCell.Font.ColorIndex = COLOR_DIFF_CELL
                    Case 2  '// セルを着色
                      targetCell.Interior.ColorIndex = COLOR_DIFF_CELL
                    Case 3  '// 枠を着色
                      targetCell.Borders.LineStyle = xlContinuous
                      targetCell.Borders.ColorIndex = COLOR_DIFF_CELL
                    Case 4  '// 該当セルを含む行以外を非表示
                      '// 将来機能
                End Select
            End If
            
            idxCellSrch = idxCellSrch + 1
            If idxCellSrch Mod 1000 = 0 Then
                Application.StatusBar = "検索中... [ " & wkSheet.Name & " " & CStr(CInt(idxCellSrch / numCellCnt)) & " ]"
            End If
        Next
    End If
    
    '// 式を検索 //////////
    If ckbSearchFormula.Value And pfGetCellCount(wkSheet.UsedRange, xlCellTypeFormulas) > -1 Then
        For Each targetCell In wkSheet.UsedRange.SpecialCells(xlCellTypeFormulas)
            If regExp.test(targetCell.FormulaLocal) Then
                Call psSetMatchedRec(wkSheet, targetCell.Row, targetCell.Column, targetCell.FormulaLocal, "数式")
                
                '// セル着色など
                Select Case cmbOutput.Value
                  Case 0  '// 何もしない
                  Case 1  '// 文字を着色
                    targetCell.Font.ColorIndex = COLOR_DIFF_CELL
                  Case 2  '// セルを着色
                    targetCell.Interior.ColorIndex = COLOR_DIFF_CELL
                  Case 3  '// 枠を着色
                    targetCell.Borders.LineStyle = xlContinuous
                    targetCell.Borders.ColorIndex = COLOR_DIFF_CELL
                  Case 4  '// 該当セルを含む行以外を非表示
                End Select
            End If
            
            idxCellSrch = idxCellSrch + 1
            If idxCellSrch Mod 1000 = 0 Then
                Application.StatusBar = "検索中... [ " & wkSheet.Name & " " & CStr(CInt(idxCellSrch / numCellCnt)) & " ]"
            End If
        Next
    End If
  
    '// シェイプ内の文字列を検索 //////////
    If ckbSearchShape.Value Then
        For Each shapeObj In wkSheet.Shapes
            If shapeObj.Type <> msoComment Then '// シェイプのうちコメントについてはコメント自体を検索するため除外
                Call psExecSearch_Shape(regExp, wkSheet, shapeObj, False)
            End If
        Next
    End If
  
    '// コメント内の文字列を検索 //////////
    If ckbSearchComment.Value Then
        For Each commentObj In wkSheet.Comments
            If regExp.test(commentObj.Text) Then
                Call psSetMatchedRec(wkSheet, commentObj.Parent.Cells.Row, commentObj.Parent.Cells.Column, commentObj.Text, "コメント")
            End If
        Next
    End If
  
    '// セル名称を検索 //////////
    '// 無効なNameがある場合のエラーを回避するため、判定ロジックを外だし（pfCheckRangeName）
    If ckbSearchName.Value Then
        For Each rangeName In wkSheet.Parent.Names  '// ブックのNamesプロパティを参照する必要がある（原因不明）
            If pfCheckRangeName(rangeName, wkSheet) Then
                If regExp.test(rangeName.Name) Then
                    Call psSetMatchedRec(wkSheet, rangeName.RefersToRange.Row, rangeName.RefersToRange.Column, rangeName.Name, "セル名称")
                End If
            End If
        Next
    End If
  
    '// ハイパーリンク先を検索 //////////
    If ckbSearchLink.Value Then
        For Each hLink In wkSheet.Hyperlinks
            If regExp.test(hLink.Address) Or regExp.test(hLink.SubAddress) Then
                Call psSetMatchedRec(wkSheet, hLink.Range.Row, hLink.Range.Column, hLink.Address & "[" & hLink.SubAddress & "]", "ハイパーリンク")
            End If
        Next
    End If
  
  '// シート名を検索 //////////
    If ckbSearchSheetName.Value Then
        If regExp.test(wkSheet.Name) Then
            Call psSetMatchedRec(wkSheet, 1, 1, wkSheet.Name, "シート名")
        End If
    End If
  
  
    '// ヘッダとフッタの文字列を検索 //////////
    If ckbSearchHeader.Value Then
        If regExp.test(wkSheet.PageSetup.LeftHeader) Then
            Call psSetMatchedRec(wkSheet, 1, 1, wkSheet.PageSetup.LeftHeader, "ヘッダ（左）")
        End If
        If regExp.test(wkSheet.PageSetup.CenterHeader) Then
            Call psSetMatchedRec(wkSheet, 1, 1, wkSheet.PageSetup.CenterHeader, "ヘッダ（中央）")
        End If
        If regExp.test(wkSheet.PageSetup.RightHeader) Then
            Call psSetMatchedRec(wkSheet, 1, 1, wkSheet.PageSetup.RightHeader, "ヘッダ（右）")
        End If
        If regExp.test(wkSheet.PageSetup.LeftFooter) Then
            Call psSetMatchedRec(wkSheet, 1, 1, wkSheet.PageSetup.LeftFooter, "フッタ（左）")
        End If
        If regExp.test(wkSheet.PageSetup.CenterFooter) Then
            Call psSetMatchedRec(wkSheet, 1, 1, wkSheet.PageSetup.CenterFooter, "フッタ（中央）")
        End If
        If regExp.test(wkSheet.PageSetup.RightFooter) Then
            Call psSetMatchedRec(wkSheet, 1, 1, wkSheet.PageSetup.RightFooter, "フッタ（右）")
        End If
    End If
  
    '// グラフを検索 //////////
    If ckbSearchGraph.Value Then
        For idxChart = 1 To wkSheet.ChartObjects.Count  '// チャートの配列は１から開始
            Set chartObj = wkSheet.ChartObjects(idxChart).Chart
            If regExp.test(pfGetChartTitle(chartObj)) Then
                Call psSetMatchedRec(wkSheet, 1, 1, chartObj.ChartTitle.Characters.Text, "チャートタイトル")
            End If
            
            For Each seriesObj In chartObj.SeriesCollection
                If regExp.test(seriesObj.Name) Then
                    Call psSetMatchedRec(wkSheet, 1, 1, seriesObj.Name, "チャート系列名")
                End If
            Next
        Next
    End If
    
    Set regExp = Nothing
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// メソッド：   シェイプ内テキスト取得
'// 説明：       シェイプ内のテキストを取得する。Charactersメソッドをサポートしない場合は例外処理でハンドリング
'//              psExecSearch_Shapeで特定されたシェイプ内のテキストを戻す
'// 引数：       shapeObj: 対象シェイプオブジェクト
'// 戻り値：     シェイプ内のテキスト。シェイプがテキストをサポートしていない場合は一律でブランク
'// ////////////////////////////////////////////////////////////////////////////
Private Function pfGetShapeText(shapeObj As Shape) As String
On Error GoTo ErrorHandler
    If shapeObj.Type = msoTextEffect Then '// ワードアートの場合
        pfGetShapeText = shapeObj.TextEffect.Text
    Else
        pfGetShapeText = shapeObj.TextFrame.Characters.Text
    End If
Exit Function

ErrorHandler:
    pfGetShapeText = BLANK
End Function


'// ////////////////////////////////////////////////////////////////////////////
'// メソッド：   検索：シェイプ
'// 説明：       シェイプ内の文字列を検索する。グループ化されている場合は再帰検索を行う。
'// 引数：       regExp: 正規表現オブジェクト
'//              wkSheet: 対象シート
'//              shapeObj: 対象シェイプオブジェクト
'//              isGrouped: グループ内オブジェクトか否か（再帰呼び出しされているか）
'// 戻り値：     なし
'// ////////////////////////////////////////////////////////////////////////////
Private Sub psExecSearch_Shape(regExp As Object, wkSheet As Worksheet, shapeObj As Shape, isGrouped As Boolean)
    Dim bffText   As String
    Dim subShape  As Shape
    
    If shapeObj.Type = msoGroup Then
        For Each subShape In shapeObj.GroupItems
            Call psExecSearch_Shape(regExp, wkSheet, subShape, True)
        Next
    Else
        bffText = pfGetShapeText(shapeObj)
        If bffText <> BLANK Then
            If regExp.test(bffText) Then
                Call psSetMatchedRec(wkSheet, IIf(isGrouped, -1, shapeObj.TopLeftCell.Row), IIf(isGrouped, -1, shapeObj.TopLeftCell.Column), bffText, "シェイプ：" & shapeObj.Name)
            End If
        End If
    End If
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// メソッド：   検索結果出力
'// 説明：       検索結果を別ブックで出力する
'// 引数：       なし
'// 戻り値：     なし
'// ////////////////////////////////////////////////////////////////////////////
Private Sub psShowResult()
    Dim wkSheet   As Worksheet
    Dim idxRow    As Long
    
    '// 出力先の設定
    With Workbooks.Add
        Set wkSheet = .ActiveSheet
    End With
  
    '// ヘッダの設定
    Call gsDrawResultHeader(wkSheet, HDR_SEARCH, 1)
  
    wkSheet.Cells.NumberFormat = "@"
    
    '// 値の設定
    For idxRow = 0 To UBound(pMatched) - 1
        wkSheet.Cells(idxRow + 2, 1).Value = pMatched(idxRow).FileName
        wkSheet.Cells(idxRow + 2, 2).Value = pMatched(idxRow).SheetName
        wkSheet.Cells(idxRow + 2, 3).Value = IIf(pMatched(idxRow).Row > 0, mdlCommon.gfGetColIndexString(pMatched(idxRow).Col) & CStr(pMatched(idxRow).Row), BLANK)
        wkSheet.Cells(idxRow + 2, 4).Value = pMatched(idxRow).TargetText
        wkSheet.Cells(idxRow + 2, 5).Value = pMatched(idxRow).NoteText
        
        If pMatched(idxRow).SavedFile And pMatched(idxRow).Row > 0 Then '// セーブされているときのみリンク設定
            ActiveSheet.Hyperlinks.Add Anchor:=wkSheet.Cells(idxRow + 2, 3), Address:=wkSheet.Cells(idxRow + 2, 1).Value, SubAddress:="'" & wkSheet.Cells(idxRow + 2, 2).Value & "'!" & wkSheet.Cells(idxRow + 2, 3).Value
        End If
    Next
  
    '// //////////////////////////////////////////////////////
    '// 書式の設定
    '// 幅の設定
    wkSheet.Columns("A:C").ColumnWidth = 10
    wkSheet.Columns("D:E").ColumnWidth = 30
    
    '// 枠線の設定
    Call gsPageSetup_Lines(wkSheet, 1)
    
'    Call wkSheet.Range(wkSheet.Cells(1, 1), wkSheet.Cells(UBound(pMatched) + 1, 5)).Select
'    Call mdlCommon.gsDrawLine_Data
'
'    '// ヘッダの修飾
'    Call wkSheet.Range("A1:E1").Select
'    Call mdlCommon.gsDrawLine_Header
    
    '//フォント
    wkSheet.Cells.Font.Name = APP_FONT
    wkSheet.Cells.Font.Size = APP_FONT_SIZE
    
    Call wkSheet.Cells(1, 1).Select
    
    '// 後処理
    Call wkSheet.Cells(1, 1).Select
    ActiveWorkbook.Saved = True    '// 閉じるときに保存を求めない
    Call gsResumeAppEvents
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// メソッド：   検索ヒットレコード登録
'// 説明：       検索にヒットした内容を配列に登録する
'// 引数：       wkSheet: 対象ワークシート
'//              Row: ヒットした行
'//              Col: ヒットした列
'//              TargetText: ヒットした値
'//              NoteText: 備考
'// 戻り値：     なし
'// ////////////////////////////////////////////////////////////////////////////
Private Sub psSetMatchedRec(wkSheet As Worksheet, Row As Long, Col As Integer, TargetText As String, NoteText As String)
    ReDim Preserve pMatched(pMatchCnt + 1)
    
    pMatched(pMatchCnt).FileName = wkSheet.Parent.Path & "\" & wkSheet.Parent.Name
    pMatched(pMatchCnt).SheetName = wkSheet.Name
    pMatched(pMatchCnt).Row = Row
    pMatched(pMatchCnt).Col = Col
    pMatched(pMatchCnt).TargetText = TargetText
    pMatched(pMatchCnt).NoteText = NoteText
    pMatched(pMatchCnt).SavedFile = IIf(wkSheet.Parent.Path = BLANK, False, True)
    
    pMatchCnt = pMatchCnt + 1
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// メソッド：   セル範囲カウント取得
'// 説明：       SpecialCells の結果カウント数を取得する
'// 引数：       targetRange: 対象範囲
'//              cellType: 取得タイプ
'// 戻り値：     範囲内の対象セル数。セルがゼロの場合は -1 を返す
'// ////////////////////////////////////////////////////////////////////////////
Private Function pfGetCellCount(targetRange As Range, cellType As Long) As Double
On Error GoTo ErrorHandler
    pfGetCellCount = targetRange.SpecialCells(cellType).Count
    Exit Function

ErrorHandler:
    pfGetCellCount = -1
End Function


'// ////////////////////////////////////////////////////////////////////////////
'// メソッド：   検索文字列の妥当性判定
'// 説明：       指定された検索文字列が正規表現として妥当か（エラーが発生しないか）を確認する
'// 引数：       regExp: 正規表現オブジェクト
'// 戻り値：     検索の成否
'// ////////////////////////////////////////////////////////////////////////////
Private Function pfCheckRegExp(regExp As Object) As Boolean
On Error GoTo ErrorHandler
    pfCheckRegExp = regExp.test(BLANK)
    pfCheckRegExp = True
    Exit Function

ErrorHandler:
    pfCheckRegExp = False
End Function


'// ////////////////////////////////////////////////////////////////////////////
'// メソッド：   セル名称の妥当性判定
'// 説明：       指定されたセル名称がwkSheetに含まれているか、および有効な名称であるかを判定する
'// 引数：       rangeName: 対象となるセル名称オブジェクト
'//              wkSheet: 対象となるシート
'// 戻り値：     妥当性の成否
'// ////////////////////////////////////////////////////////////////////////////
Private Function pfCheckRangeName(rangeName As Name, wkSheet As Worksheet) As Boolean
On Error GoTo ErrorHandler
    pfCheckRangeName = (rangeName.RefersToRange.Worksheet.Name = wkSheet.Name)
    Exit Function

ErrorHandler:
    pfCheckRangeName = False
End Function


'// ////////////////////////////////////////////////////////////////////////////
'// メソッド：   チャートタイトル取得
'// 説明：       指定されたチャートタイトルのcharactersを返す。
'// 引数：       chartObj: 対象となるチャートオブジェクト
'// 戻り値：     チャートのタイトル文字列。取得不可の場合は空白文字列
'// ////////////////////////////////////////////////////////////////////////////
Private Function pfGetChartTitle(chartObj As Chart) As String
On Error GoTo ErrorHandler
    pfGetChartTitle = chartObj.ChartTitle.Characters.Text
    Exit Function

ErrorHandler:
    pfGetChartTitle = BLANK
End Function


'// ////////////////////////////////////////////////////////////////////////////
'// END
'// ////////////////////////////////////////////////////////////////////////////
