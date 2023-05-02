VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCompSheet 
   Caption         =   "シート比較"
   ClientHeight    =   4905
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6600
   OleObjectBlob   =   "frmCompSheet.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmCompSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'// ////////////////////////////////////////////////////////////////////////////
'// プロジェクト   : 拡張ツール
'// タイトル       : シート比較フォーム
'// モジュール     : frmCompSheet
'// 説明           : シートの比較を行う
'//                  Excel2013から標準機能で実装されたため、この機能のメンテナンスは今後行わない。
'// ////////////////////////////////////////////////////////////////////////////
'// Copyright (c) by Koichiro.
'// ////////////////////////////////////////////////////////////////////////////
Option Explicit
Option Base 0

'// ////////////////////////////////////////////////////////////////////////////
'// プライベート定数
Private Const FLG_INS_ROW               As String = "$ins_extools"
Private Const FLG_DEL_ROW               As String = "$del_extools"
Private Const CLR_DIFF_INS_ROW          As Integer = 34  '// 42
Private Const CLR_DIFF_DEL_ROW          As Integer = 15  '// 48
Private Const STATINTERVAL              As Long = 100    '// ステータスバーの更新インターバル(単位:行)

'// ////////////////////////////////////////////////////////////////////////////
'// プライベート変数
'// 差分タイプ
Private Type udDiff
  sheet As String
  Row   As Integer
  Col   As Integer
  val_1 As String
  val_2 As String
  note  As String
End Type

'// 行適合タイプ
Private Type udRowPair
  row1  As Long
  row2  As Long
End Type

Private pDiff()                 As udDiff       '// 差分の結果
Private pMatched()              As udRowPair    '// 適合行の一覧


'// //////////////////////////////////////////////////////////////////
'// イベント： フォーム アクティブ時
Private Sub UserForm_Activate()
    '// 事前チェック（シート有無）
    If Not gfPreCheck() Then
        Call Me.Hide
        Exit Sub
    End If
    
  Call psSetSheetCombo(cmbSheet_1.Name)
  Call psSetSheetCombo(cmbSheet_2.Name)
End Sub


'// //////////////////////////////////////////////////////////////////
'// イベント： フォーム 初期化時
Private Sub UserForm_Initialize()
  '// コンボボックス設定
  Call gsSetCombo(cmbResultType, CMB_CMP_MARKER, 0)
  Call gsSetCombo(cmbCompareMode, CMB_CMP_METHOD, 0)
  Call gsSetCombo(cmbOutput, CMB_CMP_OUTPUT, 0)
  
  '// キャプション設定
  frmCompSheet.Caption = LBL_CMP_FORM
  cmdExecute.Caption = LBL_COM_EXEC
  cmdClose.Caption = LBL_COM_CLOSE
  cmdFile_1.Caption = LBL_COM_BROWSE
  cmdFile_2.Caption = LBL_COM_BROWSE
  mpgTarget.Pages(0).Caption = LBL_CMP_MODE_SHEET
  mpgTarget.Pages(1).Caption = LBL_CMP_MODE_BOOK
  ckbShowComments.Caption = LBL_CMP_SHOW_COMMENT
  fraOption.Caption = LBL_CMP_OPTIONS
  lblOriginalSheet.Caption = LBL_CMP_SHEET1
  lblTargetSheet.Caption = LBL_CMP_SHEET2
  lblOriginalBook.Caption = LBL_CMP_BOOK1
  lblTargetBook.Caption = LBL_CMP_BOOK2
  lblOutput.Caption = LBL_CMP_RESULT
  lblMarker.Caption = LBL_CMP_MARKER
  lblMethod.Caption = LBL_CMP_METHOD
End Sub


'// //////////////////////////////////////////////////////////////////
'// イベント： 閉じるボタン クリック時
Private Sub cmdClose_Click()
  Call Me.Hide
End Sub


'// //////////////////////////////////////////////////////////////////
'// イベント： 参照ボタン クリック時
Private Sub cmdFile_1_Click()
  txtFileName_1.Text = pfGetFileName(txtFileName_1.Text)
End Sub


'// //////////////////////////////////////////////////////////////////
'// イベント： 参照ボタン クリック時
Private Sub cmdFile_2_Click()
  txtFileName_2.Text = pfGetFileName(txtFileName_2.Text)
End Sub


'// //////////////////////////////////////////////////////////////////
'// イベント： 実行ボタン クリック時
Private Sub cmdExecute_Click()
  Select Case mpgTarget.Value
    Case 0
      Call psCompSheet
    Case 1
      Call psCompBook
  End Select
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// メソッド：   シート比較
'// 説明：       シートの比較を行う
'// 引数：       なし
'// ////////////////////////////////////////////////////////////////////////////
Private Sub psCompSheet()
On Error GoTo ErrorHandler
  Dim errCnt    As Long   '// エラー数
  
  Application.ScreenUpdating = False
  '// 比較実行
  Erase pDiff
  Call psExecComp(Worksheets(cmbSheet_1.Text), Worksheets(cmbSheet_2.Text), cmbResultType.Value, cmbCompareMode.Value, errCnt)
    
  '// 結果の出力
  If errCnt > 0 Then
    Call psShowResult(ActiveWorkbook)
  Else
    '// ゼロ件の場合はメッセージを表示
    Call MsgBox(MSG_NO_DIFF, vbOKOnly, APP_TITLE)
  End If
  
  Call Me.Hide
  Application.ScreenUpdating = True
  Exit Sub
  
ErrorHandler:
  Call gsShowErrorMsgDlg("frmCompSheet.psCompSheet", Err)
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// メソッド：   ブック比較
'// 説明：       ブックの比較を行う
'// 引数：       なし
'// ////////////////////////////////////////////////////////////////////////////
Private Sub psCompBook()
On Error GoTo ErrorHandler
  Dim errCnt    As Long       '// エラー数
  Dim idx       As Integer    '// シートのインデクス
  Dim book1     As Workbook   '// 対象ブック（１）
  Dim book2     As Workbook   '// 対象ブック（２）
  Dim cntRows   As Long
  Dim cntCols   As Integer
  
  '// 入力チェック
  If (Trim(txtFileName_1.Text) = BLANK) Or (Trim(txtFileName_2.Text) = BLANK) Then
    Call MsgBox(MSG_ERROR_NEED_BOOKNAME, vbOKOnly, APP_TITLE)
    If Trim(txtFileName_1.Text) = BLANK Then
      Call txtFileName_1.SetFocus
    Else
      Call txtFileName_2.SetFocus
    End If
    Exit Sub
  End If

  Application.ScreenUpdating = False
  '// シートを開く
  Set book1 = pfGetBook(txtFileName_1.Text)
  Set book2 = pfGetBook(txtFileName_2.Text)
  
  If (book1 Is Nothing) Or (book2 Is Nothing) Then
    Application.ScreenUpdating = True
    Exit Sub
  End If
  
  '// 比較実行（シート構成）
  Erase pDiff
  errCnt = 0
  For idx = 1 To book1.Worksheets.Count
    '// 配列への格納
    ReDim Preserve pDiff(errCnt + 1)
    pDiff(errCnt).val_1 = "シート： " & book1.Worksheets(idx).Name
    If idx <= book2.Worksheets.Count Then
      pDiff(errCnt).val_2 = "シート： " & book2.Worksheets(idx).Name
    End If
    '// シート名が違う場合
    If pDiff(errCnt).val_1 <> pDiff(errCnt).val_2 Then
      pDiff(errCnt).sheet = book2.Worksheets(errCnt + 1).Name
      pDiff(errCnt).Row = 1
      pDiff(errCnt).Col = 1
      pDiff(errCnt).note = MSG_SHEET_NAME
      errCnt = errCnt + 1
    End If
  Next
  
  '// シート構成が異なる場合には終了
  If errCnt > 0 Then
    Call MsgBox(MSG_UNMATCH_SHEET, vbOKOnly, APP_TITLE)
    Call psShowResult(book2)
    Call Me.Hide
    Set book1 = Nothing
    Set book2 = Nothing
    Application.ScreenUpdating = True
    Exit Sub
  End If
  
  '// 比較実行（シート毎）
  Erase pDiff
  errCnt = 0
  For idx = 1 To book1.Worksheets.Count
    Call psExecComp(book1.Worksheets(idx), book2.Worksheets(idx), cmbResultType.Value, cmbCompareMode.Value, errCnt)
  Next
  
  If errCnt > 0 Then
    '// 結果の出力
    Call psShowResult(book2)
    Call Me.Hide
  Else
    '// ゼロ件の場合はメッセージを表示
    Call MsgBox(MSG_NO_DIFF, vbOKOnly, APP_TITLE)
  End If
  
  Set book1 = Nothing
  Set book2 = Nothing
  Application.ScreenUpdating = True
  Exit Sub
  
ErrorHandler:
  Call gsShowErrorMsgDlg("frmCompBook.psCompBook", Err)
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// メソッド：   シートのコンボ設定
'// 説明：       指定されたブックに含まれるシートを検索し、コンボボックスに設定する
'// 引数：       comboName: コンボボックス名称
'// ////////////////////////////////////////////////////////////////////////////
Private Sub psSetSheetCombo(comboName As String)
  Dim combo         As ComboBox
  Dim wkSheet       As Worksheet
  Dim currentSheet  As String  '// コンボボックス初期化前のシート名
  Dim defaultIdx    As Integer
  
  '// 初期化
  defaultIdx = 0
  Set combo = Me.Controls(comboName)
  currentSheet = combo.Text
  Call combo.Clear
  
  '// シートを取得
  For Each wkSheet In ActiveWorkbook.Worksheets
    Call combo.AddItem(wkSheet.Name)
    If wkSheet.Name = currentSheet Then
      defaultIdx = combo.ListCount - 1
    End If
  Next
  
  combo.ListIndex = defaultIdx
  Set combo = Nothing
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// メソッド：   ファイル名取得
'// 説明：       ダイアログを表示し、ファイル名を返す。
'// 引数：       defaultVal: ファイル名
'// 戻り値：     フルパスのファイル名
'// ////////////////////////////////////////////////////////////////////////////
Private Function pfGetFileName(defaultVal As String)
'  Call gsCheckVersion
  
  pfGetFileName = Application.GetOpenFilename(FileFilter:=Replace(APP_EXL_FILE, "#", EXCEL_FILE_EXT))
  If pfGetFileName = False Then
    pfGetFileName = defaultVal
  End If
End Function


'// ////////////////////////////////////////////////////////////////////////////
'// メソッド：   ブックのオープン
'// 説明：       ブックオブジェクトを返す
'// 引数：       fileName: ファイル名
'// 戻り値：     ブックオブジェクト
'// ////////////////////////////////////////////////////////////////////////////
Private Function pfGetBook(FileName As String) As Workbook
On Error GoTo ErrorHandler
  Set pfGetBook = Workbooks.Open(FileName:=FileName, ReadOnly:=False)
  Exit Function

ErrorHandler:
  Call MsgBox(MSG_NO_FILE & " [" & FileName & " ]", vbOKOnly, APP_TITLE)
  Set pfGetBook = Nothing
End Function


'// ////////////////////////////////////////////////////////////////////////////
'// メソッド：   比較
'// 説明：       比較を行う
'// 引数：       wkSheet1, wkSheet2: 比較対象シート
'//              compResult 出力形式 0:何もしない  1:文字を着色  2:セルを着色
'//              compMode 比較モード 0:テキスト  1:値  2:値またはテキスト
'//              cnt: 差分数（呼び出し元引継ぎ）
'// 戻り値：     なし
'// ////////////////////////////////////////////////////////////////////////////
Private Sub psExecComp(wkSheet1 As Worksheet, wkSheet2 As Worksheet, _
                       compResult As Integer, compMode As Integer, ByRef cnt As Long)
On Error GoTo ErrorHandler
  Dim idxRow          As Long           '// 抽出時の行インデクス
  Dim idxCol          As Integer        '// 抽出時の列インデクス
  Dim tRange          As udTargetRange  '// 抽出範囲
  Dim isDiff_v        As Boolean        '// 値の違い有無判定
  Dim isDiff_t        As Boolean        '// テキストの違い有無判定
  Dim isDiff_f        As Boolean        '// 書式の違い有無判定
  Dim isDiff          As Boolean        '// 差分有無の総合判定
  
  '// 検索範囲の設定
  tRange.minRow = IIf(wkSheet1.UsedRange.Row < wkSheet2.UsedRange.Row, wkSheet1.UsedRange.Row, wkSheet2.UsedRange.Row)
  tRange.minCol = IIf(wkSheet1.UsedRange.Column < wkSheet2.UsedRange.Column, wkSheet1.UsedRange.Column, wkSheet2.UsedRange.Column)
  tRange.maxRow = IIf((wkSheet1.UsedRange.Row + wkSheet1.UsedRange.Rows.Count) > (wkSheet2.UsedRange.Row + wkSheet2.UsedRange.Rows.Count), (wkSheet1.UsedRange.Row + wkSheet1.UsedRange.Rows.Count - 1), (wkSheet2.UsedRange.Row + wkSheet2.UsedRange.Rows.Count - 1))
  tRange.maxCol = IIf((wkSheet1.UsedRange.Column + wkSheet1.UsedRange.Columns.Count) > (wkSheet2.UsedRange.Column + wkSheet2.UsedRange.Columns.Count), (wkSheet1.UsedRange.Column + wkSheet1.UsedRange.Columns.Count - 1), (wkSheet2.UsedRange.Column + wkSheet2.UsedRange.Columns.Count - 1))
  
  '// 各シートが２行以上、かつ合計５行以上の場合のみ、行差分の精査を実施（行が少ない場合は精査での例外処理が面倒なため）
  If wkSheet1.UsedRange.Rows.Count > 1 And wkSheet2.UsedRange.Rows.Count > 1 And wkSheet1.UsedRange.Rows.Count + wkSheet2.UsedRange.Rows.Count > 4 Then
    Call psPadRowDiff(compMode, wkSheet1, wkSheet2, tRange.maxRow, tRange.maxCol)
  End If
  
  '// 検索範囲の設定
  tRange.maxRow = IIf((wkSheet1.UsedRange.Row + wkSheet1.UsedRange.Rows.Count) > (wkSheet2.UsedRange.Row + wkSheet2.UsedRange.Rows.Count), (wkSheet1.UsedRange.Row + wkSheet1.UsedRange.Rows.Count - 1), (wkSheet2.UsedRange.Row + wkSheet2.UsedRange.Rows.Count - 1))
  
  
  '// 比較
  Call wkSheet2.Activate
  For idxRow = tRange.minRow To tRange.maxRow
    If wkSheet2.Cells(idxRow, 1).NoteText = FLG_INS_ROW Then
      wkSheet2.Cells(idxRow, 1).NoteText (MSG_INS_ROW)
      '// 配列への格納
      ReDim Preserve pDiff(cnt + 1)
      pDiff(cnt).sheet = wkSheet2.Name
      pDiff(cnt).Row = idxRow
      pDiff(cnt).Col = 1
      pDiff(cnt).note = MSG_INS_ROW
      cnt = cnt + 1
    ElseIf wkSheet2.Cells(idxRow, 1).NoteText = FLG_DEL_ROW Then
      wkSheet2.Cells(idxRow, 1).NoteText (MSG_DEL_ROW)
      '// 配列への格納
      ReDim Preserve pDiff(cnt + 1)
      pDiff(cnt).sheet = wkSheet2.Name
      pDiff(cnt).Row = idxRow
      pDiff(cnt).Col = 1
      pDiff(cnt).note = MSG_DEL_ROW
      cnt = cnt + 1
    Else
      For idxCol = tRange.minCol To tRange.maxCol
        isDiff_v = False
        isDiff_t = False
        isDiff_f = False
        '// テキストの違いを確認
        isDiff_t = (wkSheet1.Cells(idxRow, idxCol).Text <> wkSheet2.Cells(idxRow, idxCol).Text)
        '// 値の違いを確認
        If IsError(wkSheet1.Cells(idxRow, idxCol)) And IsError(wkSheet2.Cells(idxRow, idxCol)) Then
          isDiff_v = False
        ElseIf IsError(wkSheet1.Cells(idxRow, idxCol)) Xor IsError(wkSheet2.Cells(idxRow, idxCol)) Then
          isDiff_v = True
        Else
          isDiff_v = (wkSheet1.Cells(idxRow, idxCol).Value <> wkSheet2.Cells(idxRow, idxCol).Value)
        End If
        '// 書式の違いを確認
        isDiff_f = (CStr(wkSheet1.Cells(idxRow, idxCol).NumberFormat) <> CStr(wkSheet2.Cells(idxRow, idxCol).NumberFormat))
        
        isDiff = (isDiff_t And compMode = 0) Or (isDiff_v And compMode = 1) Or ((compMode = 2) And (isDiff_v And isDiff_t))
        If isDiff Then
          '// 着色（指定時）
          If Not wkSheet2.ProtectContents Then
            Select Case compResult
              Case 1   '// 文字を着色
                wkSheet2.Cells(idxRow, idxCol).Font.ColorIndex = COLOR_DIFF_CELL
              Case 2   '// セルを着色
                wkSheet2.Cells(idxRow, idxCol).Interior.ColorIndex = COLOR_DIFF_CELL
              Case 3   '// 枠を着色
                wkSheet2.Cells(idxRow, idxCol).Borders(xlEdgeLeft).LineStyle = xlContinuous
                wkSheet2.Cells(idxRow, idxCol).Borders(xlEdgeLeft).ColorIndex = COLOR_DIFF_CELL
                wkSheet2.Cells(idxRow, idxCol).Borders(xlEdgeTop).LineStyle = xlContinuous
                wkSheet2.Cells(idxRow, idxCol).Borders(xlEdgeTop).ColorIndex = COLOR_DIFF_CELL
                wkSheet2.Cells(idxRow, idxCol).Borders(xlEdgeBottom).LineStyle = xlContinuous
                wkSheet2.Cells(idxRow, idxCol).Borders(xlEdgeBottom).ColorIndex = COLOR_DIFF_CELL
                wkSheet2.Cells(idxRow, idxCol).Borders(xlEdgeRight).LineStyle = xlContinuous
                wkSheet2.Cells(idxRow, idxCol).Borders(xlEdgeRight).ColorIndex = COLOR_DIFF_CELL
            End Select
          End If
          '// コメント
          If Not wkSheet2.ProtectContents Then
            Select Case compMode
              Case 0   '// テキスト
                Call wkSheet2.Cells(idxRow, idxCol).NoteText(IIf(wkSheet1.Cells(idxRow, idxCol).Text = BLANK, "<Blank>", wkSheet1.Cells(idxRow, idxCol).Text))
              Case 1   '// 値
                If IsError(wkSheet1.Cells(idxRow, idxCol)) Then
                  Call wkSheet2.Cells(idxRow, idxCol).NoteText(wkSheet1.Cells(idxRow, idxCol).Text)
                Else
                  Call wkSheet2.Cells(idxRow, idxCol).NoteText(IIf(wkSheet1.Cells(idxRow, idxCol).Value = BLANK, "<Blank>", wkSheet1.Cells(idxRow, idxCol).Value))
                End If
              Case 2   '// テキストまたは値
                If IsError(wkSheet1.Cells(idxRow, idxCol)) Then
                  Call wkSheet2.Cells(idxRow, idxCol).NoteText(wkSheet1.Cells(idxRow, idxCol).Text)
                Else
                  Call wkSheet2.Cells(idxRow, idxCol).NoteText("【ﾃｷｽﾄ】" & wkSheet1.Cells(idxRow, idxCol).Text & vbLf & "【値】" & wkSheet1.Cells(idxRow, idxCol).Value)
                End If
            End Select
          End If
          
          '// 配列への格納
          ReDim Preserve pDiff(cnt + 1)
          pDiff(cnt).sheet = wkSheet2.Name
          pDiff(cnt).Row = idxRow
          pDiff(cnt).Col = idxCol
          '// 値が違う場合
          If isDiff_v Or isDiff_t Then
            pDiff(cnt).val_1 = IIf(IsError(wkSheet1.Cells(idxRow, idxCol)), wkSheet1.Cells(idxRow, idxCol).Text, wkSheet1.Cells(idxRow, idxCol).Value)
            pDiff(cnt).val_2 = IIf(IsError(wkSheet2.Cells(idxRow, idxCol)), wkSheet2.Cells(idxRow, idxCol).Text, wkSheet2.Cells(idxRow, idxCol).Value)
          End If
          
          '// 書式が違う場合
          If isDiff_f Then
            pDiff(cnt).note = "書式"
          End If
          
          cnt = cnt + 1
          '// キー割込
          If GetAsyncKeyState(27) <> 0 Then
            Application.StatusBar = False
            Exit Sub
          End If
        End If
      Next
    End If
    
    '// ステータスバー更新
    If idxRow Mod STATINTERVAL = 0 Then
      Application.StatusBar = "セル内容比較中... [ 行: " & CStr(idxRow) & " / 差分: " & CStr(cnt) & " ]" & IIf(wkSheet1.Name = wkSheet2.Name, wkSheet1.Name, BLANK)
    End If
  Next
  
  Application.StatusBar = False
  Exit Sub
ErrorHandler:
  Call gsShowErrorMsgDlg("frmCompSheet.psExecComp", Err)
  Application.StatusBar = False
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// メソッド：   比較結果出力
'// 説明：       比較結果を別ブックで出力する
'// 引数：       なし
'// 戻り値：     なし
'// 修正履歴：   なし
'// ////////////////////////////////////////////////////////////////////////////
Private Sub psShowResult(wkBook As Workbook)
  Dim wkSheet   As Worksheet
  Dim idxRow    As Integer
  Dim needLink  As Boolean
  
  needLink = Not wkBook.MultiUserEditing And (cmbOutput.Value <> "0")
  
  '// 出力先の設定
  Select Case cmbOutput.Value
    Case "0"
      Call Workbooks.Add
      Set wkSheet = ActiveWorkbook.ActiveSheet
    Case "1"
      Set wkSheet = wkBook.Sheets.Add(After:=wkBook.Sheets(wkBook.Sheets.Count))
  End Select
  
  '// ヘッダの設定
  wkSheet.Cells(1, 1).Value = "シート"
  wkSheet.Cells(1, 2).Value = "セル"
  wkSheet.Cells(1, 3).Value = "比較もとの値 (" & IIf(mpgTarget.Value = 0, cmbSheet_1.Text, txtFileName_1.Text) & ")"
  wkSheet.Cells(1, 4).Value = "比較先の値 (" & IIf(mpgTarget.Value = 0, cmbSheet_2.Text, txtFileName_2.Text) & ")"
  wkSheet.Cells(1, 5).Value = "備考"
  
  wkSheet.Columns("C:D").NumberFormat = "@"
  
  '// 差分の設定
  For idxRow = 0 To UBound(pDiff) - 1
    wkSheet.Cells(idxRow + 2, 1).Value = pDiff(idxRow).sheet
    wkSheet.Cells(idxRow + 2, 2).Value = mdlCommon.gfGetColIndexString(pDiff(idxRow).Col) & CStr(pDiff(idxRow).Row)
    If needLink Then
      Call wkSheet.Cells(idxRow + 2, 2).Hyperlinks.Add(Anchor:=Cells(idxRow + 2, 2), Address:=BLANK, SubAddress:="'" & pDiff(idxRow).sheet & "'!" & wkSheet.Cells(idxRow + 2, 2).Value)
    End If
    wkSheet.Cells(idxRow + 2, 3).Value = pDiff(idxRow).val_1
    wkSheet.Cells(idxRow + 2, 4).Value = pDiff(idxRow).val_2
    wkSheet.Cells(idxRow + 2, 5).Value = pDiff(idxRow).note
  Next
  
  '// //////////////////////////////////////////////////////
  '// 書式の設定
  '// 幅の設定
  wkSheet.Columns("A").ColumnWidth = 10
  wkSheet.Columns("B").ColumnWidth = 8
  wkSheet.Columns("C:E").ColumnWidth = 30
  
  '// 枠線の設定
  Call wkSheet.Range(wkSheet.Cells(1, 1), wkSheet.Cells(UBound(pDiff) + 1, 5)).Select
  Call gsDrawLine_Data
  
  '// ヘッダの修飾
  Call wkSheet.Range("A1:E1").Select
  Call gsDrawLine_Header
  
  '//フォント
  wkSheet.Cells.Select
  Selection.Font.Name = APP_FONT
  Selection.Font.Size = APP_FONT_SIZE
  Call wkSheet.Cells(1, 1).Select
  
  '// 閉じるときに保存を求めない
  If cmbOutput.Value = "0" Then
    ActiveWorkbook.Saved = True
  End If
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// メソッド：   行移動（同一行である場合に下位行へ移動）
'// 説明：       ＯＮＤのサブメソッド
'// 引数：       compMode 比較モード 0:テキスト  1:値  2:値またはテキスト
'//              wkSheet1, wkSheet2: 比較対象シート
'//              idxRow1, idxRow2 対象シート内の比較対象行
'//              maxCol: 比較列数
'// 戻り値：     シート２の一致行番号
'// ////////////////////////////////////////////////////////////////////////////
Private Function pfSnake(compMode As Integer, wkSheet1 As Worksheet, wkSheet2 As Worksheet, _
                         idxRow1 As Long, idxRow2 As Long, _
                         maxRow As Long, maxCol As Integer) As Long
  If idxRow1 < 1 Or idxRow2 < 1 Then
    pfSnake = idxRow2
    Exit Function
  End If
  
  Do While (idxRow1 < maxRow) And (idxRow2 < maxRow) And (pfGetRowScore(compMode, wkSheet1, wkSheet2, idxRow1, idxRow2, maxCol, False) > 0)

    
    ReDim Preserve pMatched(UBound(pMatched) + 1)
    pMatched(UBound(pMatched)).row1 = idxRow1
    pMatched(UBound(pMatched)).row2 = idxRow2
    
    idxRow1 = idxRow1 + 1
    idxRow2 = idxRow2 + 1
    
'    If (idxRow1 Mod STATINTERVAL = 0) Or (idxRow2 Mod STATINTERVAL = 0) Then
'      Application.StatusBar = "行差分分析中... [ 比較元: " & CStr(idxRow1) & " / 比較先: " & CStr(idxRow2) & " ]" & IIf(wkSheet1.Name = wkSheet2.Name, wkSheet1.Name, BLANK)
'    End If
  Loop
  
  Application.StatusBar = False
  pfSnake = idxRow2
End Function


'// ////////////////////////////////////////////////////////////////////////////
'// メソッド：   行比較
'// 説明：       指定行が同一であるかを真偽値で返す
'// 引数：       compMode 比較モード 0:テキスト  1:値  2:値またはテキスト
'//              wkSheet1, wkSheet2: 比較対象シート
'//              idxRow1, idxRow2 対象シート内の比較対象行
'//              maxCol: 比較列数
'//              getScore: true:スコアを取得する, false:真偽値として 0 または1を返す。
'// 戻り値：     0～1の実数。 0は行が完全に一致しない。1は行が完全に一致する
'// ////////////////////////////////////////////////////////////////////////////
Private Function pfGetRowScore(compMode As Integer, wkSheet1 As Worksheet, wkSheet2 As Worksheet, idxRow1 As Long, idxRow2 As Long, maxCol As Integer, getScore As Boolean) As Double
On Error GoTo ErrorHandler
  Dim idxCol    As Integer
  Dim isDiff_v  As Boolean        '// 値の違い有無判定
  Dim isDiff_t  As Boolean        '// テキストの違い有無判定
  Dim isDiff_f  As Boolean        '// 書式の違い有無判定
  Dim score     As Long           '// 将来対応：スコアによる行類似度の優劣判定
  
  For idxCol = 1 To maxCol
    isDiff_v = False
    isDiff_t = False
    isDiff_f = False
    '// テキストの違いを確認
    isDiff_t = (wkSheet1.Cells(idxRow1, idxCol).Text <> wkSheet2.Cells(idxRow2, idxCol).Text)
    '// 値の違いを確認
    isDiff_v = (wkSheet1.Cells(idxRow1, idxCol).Value <> wkSheet2.Cells(idxRow2, idxCol).Value)
    '// 書式の違いを確認
    isDiff_f = (CStr(wkSheet1.Cells(idxRow1, idxCol).NumberFormat) <> CStr(wkSheet2.Cells(idxRow2, idxCol).NumberFormat))
    
    If (isDiff_t And compMode = 0) Or (isDiff_v And compMode = 1) Or ((compMode = 2) And (isDiff_v And isDiff_t)) Then
      If Not getScore Then
        pfGetRowScore = 0
        Exit Function
      End If
    Else
      score = score + 1
    End If
  Next
  
  pfGetRowScore = CDbl(score / idxCol)
  Exit Function

ErrorHandler:
  pfGetRowScore = 0
End Function


'// ////////////////////////////////////////////////////////////////////////////
'// メソッド：   エディットグラフ走査
'// 説明：       O(ND) アルゴリズムでの走査を行う
'// 引数：       compMode 比較モード 0:テキスト  1:値  2:値またはテキスト
'//              wkSheet1, wkSheet2: 比較対象シート
'//              maxRow, maxCol: 比較列数
'// 戻り値：     true:行は同一  false:行は不一致
'// ////////////////////////////////////////////////////////////////////////////
Public Sub psOnd(compMode As Integer, wkSheet1 As Worksheet, wkSheet2 As Worksheet, _
                         maxRow As Long, maxCol As Integer)
  Dim currentIdx  As udRowPair
  Dim offset      As Long
  Dim idxSed      As Long
  Dim idxRoute    As Long
  Dim aryTemp()   As Long
  
  currentIdx.row1 = 0
  currentIdx.row2 = 0
  offset = maxRow
  
  ReDim pMatched(0)
  ReDim aryTemp(maxRow * 2)
  
  aryTemp(1 + offset) = 0
  
  For idxSed = 0 To maxRow * 2 'M + N
    For idxRoute = (-1 * idxSed) To idxSed Step 2
      If (idxRoute = -1 * idxSed) Then
        currentIdx.row2 = aryTemp(idxRoute + 1 + offset)
      ElseIf ((idxRoute <> idxSed) And aryTemp(idxRoute - 1 + offset) < aryTemp(idxRoute + 1 + offset)) Then
        currentIdx.row2 = aryTemp(idxRoute + 1 + offset)
      Else
        currentIdx.row2 = aryTemp(idxRoute - 1 + offset) + 1
      End If
      
      currentIdx.row1 = currentIdx.row2 - idxRoute
      
      aryTemp(idxRoute + offset) = pfSnake(compMode, wkSheet1, wkSheet2, currentIdx.row1, currentIdx.row2, maxRow, maxCol)
      If (currentIdx.row1 >= maxRow Or currentIdx.row2 >= maxRow) Then
        Exit Sub
      End If
    Next
  Next
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// メソッド：   行差分補正
'// 説明：       行の精査を行った後、差分行を挿入する
'// 引数：       compMode 比較モード 0:テキスト  1:値  2:値またはテキスト
'//              wkSheet1, wkSheet2: 比較対象シート
'//              maxRow, maxCol: 比較行数、列数
'// ////////////////////////////////////////////////////////////////////////////
Public Sub psPadRowDiff(compMode As Integer, wkSheet1 As Worksheet, wkSheet2 As Worksheet, _
                        maxRow As Long, maxCol As Integer)
  Dim idxMatched  As Long
  Dim pad1        As Long
  Dim pad2        As Long
  Dim rowDiff1    As Long
  Dim rowDiff2    As Long
  Dim rowDiff     As Long  '// 最後の行を補正する際の差分
  Dim cnt         As Long
  Dim rowCntLimit As Long  '// 行追加の上限値（初期状態の２シートの行数の合計）
  
  Call psOnd(compMode, wkSheet1, wkSheet2, maxRow, maxCol)
  
  rowCntLimit = wkSheet1.UsedRange.Rows.Count + wkSheet2.UsedRange.Rows.Count
  ReDim pDiffRows(0)
  idxMatched = 1
  Do
    If idxMatched > UBound(pMatched) Then
      Exit Do
    End If
    
    rowDiff1 = pMatched(idxMatched).row1 - IIf(idxMatched = 0, 0, pMatched(idxMatched - 1).row1)
    rowDiff2 = pMatched(idxMatched).row2 - IIf(idxMatched = 0, 0, pMatched(idxMatched - 1).row2)
    
    If rowDiff1 > 1 Or rowDiff2 > 1 Then
      '// シート２に行追加
      If (rowDiff2 = 1) Or ((rowDiff1 > rowDiff2) And rowDiff1 <> 1) Then
        For cnt = 1 To (pMatched(idxMatched).row1 + pad1) - (pMatched(idxMatched).row2 + pad2)
          Call wkSheet2.Rows(pMatched(idxMatched).row2 + pad2 + cnt - 1).Insert(Shift:=xlDown)
          wkSheet2.Rows(pMatched(idxMatched).row2 + pad2 + cnt - 1).Interior.ColorIndex = CLR_DIFF_DEL_ROW
          Call wkSheet1.Rows(pMatched(idxMatched).row2 + pad2 + cnt - 1).Copy
          Call wkSheet2.Rows(pMatched(idxMatched).row2 + pad2 + cnt - 1).PasteSpecial(Paste:=xlValues)
          If ROW_DIFF_STRIKETHROUGH Then
            wkSheet2.Rows(pMatched(idxMatched).row2 + pad2 + cnt - 1).Font.Strikethrough = True
          End If
          Call wkSheet2.Cells(pMatched(idxMatched).row2 + pad2 + cnt - 1, 1).NoteText(FLG_DEL_ROW) '// 後で削除
        Next
        pad2 = pad2 + (pMatched(idxMatched).row1 + pad1) - (pMatched(idxMatched).row2 + pad2)
      ElseIf (rowDiff1 = 1) Or (rowDiff1 < rowDiff2) Then
        For cnt = 1 To (pMatched(idxMatched).row2 + pad2) - (pMatched(idxMatched).row1 + pad1)
          Call wkSheet1.Rows(pMatched(idxMatched).row1 + pad1 + cnt - 1).Insert(Shift:=xlDown)
          wkSheet1.Rows(pMatched(idxMatched).row1 + pad1 + cnt - 1).Interior.ColorIndex = CLR_DIFF_DEL_ROW
          Call wkSheet2.Rows(pMatched(idxMatched).row1 + pad1 + cnt - 1).Copy
          Call wkSheet1.Rows(pMatched(idxMatched).row1 + pad1 + cnt - 1).PasteSpecial(Paste:=xlValues)
          If ROW_DIFF_STRIKETHROUGH Then
            wkSheet1.Rows(pMatched(idxMatched).row1 + pad1 + cnt - 1).Font.Strikethrough = True
          End If
          Call wkSheet2.Cells(pMatched(idxMatched).row1 + pad1 + cnt - 1, 1).NoteText(FLG_INS_ROW)  '// 後で削除
          wkSheet2.Rows(pMatched(idxMatched).row1 + pad1 + cnt - 1).Interior.ColorIndex = CLR_DIFF_INS_ROW
        Next
        pad1 = pad1 + (pMatched(idxMatched).row2 + pad2) - (pMatched(idxMatched).row1 + pad1)
      End If
    End If
    idxMatched = idxMatched + 1
    
    If wkSheet1.UsedRange.Rows.Count > rowCntLimit Or wkSheet2.UsedRange.Rows.Count > rowCntLimit Then
      Exit Do
    End If
  Loop
  
  '// 最後の行について補正
  rowDiff = (wkSheet1.UsedRange.Row + wkSheet1.UsedRange.Rows.Count) - (wkSheet2.UsedRange.Row + wkSheet2.UsedRange.Rows.Count)
  If rowDiff > 0 Then
    For cnt = wkSheet1.UsedRange.Row + wkSheet1.UsedRange.Rows.Count - rowDiff To wkSheet1.UsedRange.Row + wkSheet1.UsedRange.Rows.Count - 1
      wkSheet2.Rows(cnt).Interior.ColorIndex = CLR_DIFF_DEL_ROW
      wkSheet1.Rows(cnt).Interior.ColorIndex = CLR_DIFF_INS_ROW
      Call wkSheet1.Rows(cnt).Copy
      Call wkSheet2.Rows(cnt).PasteSpecial(Paste:=xlValues)
      If ROW_DIFF_STRIKETHROUGH Then
        wkSheet2.Rows(cnt).Font.Strikethrough = True
      End If
      Call wkSheet2.Cells(cnt, 1).NoteText(FLG_DEL_ROW) '// 後で削除
    Next
  ElseIf rowDiff < 0 Then
    For cnt = wkSheet2.UsedRange.Row + wkSheet2.UsedRange.Rows.Count + rowDiff To wkSheet2.UsedRange.Row + wkSheet2.UsedRange.Rows.Count - 1
      wkSheet2.Rows(cnt).Interior.ColorIndex = CLR_DIFF_INS_ROW
      wkSheet1.Rows(cnt).Interior.ColorIndex = CLR_DIFF_DEL_ROW
      Call wkSheet2.Rows(cnt).Copy
      Call wkSheet1.Rows(cnt).PasteSpecial(Paste:=xlValues)
      If ROW_DIFF_STRIKETHROUGH Then
        wkSheet1.Rows(cnt).Font.Strikethrough = True
      End If
      Call wkSheet2.Cells(cnt, 1).NoteText(FLG_INS_ROW) '// 後で削除
    Next
      
  End If
End Sub

'// ////////////////////////////////////////////////////////////////////////////
'// END
'// ////////////////////////////////////////////////////////////////////////////

