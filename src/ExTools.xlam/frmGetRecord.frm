VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmGetRecord 
   Caption         =   "SQL文実行"
   ClientHeight    =   9465
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8175
   OleObjectBlob   =   "frmGetRecord.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmGetRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'// ////////////////////////////////////////////////////////////////////////////
'// プロジェクト   : 拡張ツール
'// タイトル       : SQL実行
'// モジュール     : frmGetRecord
'// 説明           : SELECT スクリプトの結果をエクセルに出力する。
'// ////////////////////////////////////////////////////////////////////////////
'// Copyright (c) by Koichiro.
'// ////////////////////////////////////////////////////////////////////////////
Option Explicit
Option Base 0


'// ////////////////////////////////////////////////////////////////////////////
'// プライベート変数
Private pFileName           As String   '// ファイル名


'// //////////////////////////////////////////////////////////////////
'// イベント： 検索実行ボタン クリック時
Private Sub cmdExecute_Click()
    Call psExecSearch
End Sub


'// //////////////////////////////////////////////////////////////////
'// イベント： ログインボタン クリック時
Private Sub cmdLogin_Click()
    frmLogin.Show
End Sub


'// //////////////////////////////////////////////////////////////////
'// イベント： 閉じるボタン クリック時
Private Sub cmdClose_Click()
    Call Me.Hide
End Sub


'// //////////////////////////////////////////////////////////////////
'// イベント： フォーム 初期化時
Private Sub UserForm_Initialize()
    Call gsSetCombo(cmbDateFormat, "0,yyyy/mm/dd;1,yyyy/mm/dd hh:mm:ss", 0)
    Call gsSetCombo(cmdHeader, CMB_GRC_HEADER, 0)

    '// キャプション設定
    frmGetRecord.Caption = LBL_GRC_FORM
    fraOptions.Caption = LBL_GRC_OPTIONS
    cmdLogin.Caption = LBL_GRC_LOGIN
    cmdExecute.Caption = LBL_GRC_SEARCH
    cmdClose.Caption = LBL_COM_CLOSE
    lblDateFormat.Caption = LBL_GRC_DATE_FORMAT
    lblHeader.Caption = LBL_GRC_HEADER
    lblStatement.Caption = LBL_GRC_SCRIPT
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// メソッド：    クエリー実行
'// 説明：        引数のクエリーを実行し、シートに出力します。
'// ////////////////////////////////////////////////////////////////////////////
Private Sub psExecSearch()
On Error GoTo ErrorHandler
    Dim wkSheet       As Worksheet
    Dim rst           As Object
    Dim headerRows    As Integer
  
    If gADO Is Nothing Then
        Call frmLogin.Show
        If gADO Is Nothing Then
            Exit Sub
        End If
    End If
  
    '// メインＳＱＬの問い合わせ
    Application.StatusBar = MSG_QUERY
    Set rst = gADO.GetRecordset(txtScript.Text)
  
    If rst Is Nothing Then
        Call gsShowErrorMsgDlg("frmGetRecord.psExecSearch", Err, gADO)
        Application.StatusBar = False
        Exit Sub
    End If
  
    If rst.Fields.Count > 0 Then    '// SELECT文の場合
        If Not rst.EOF Then
            Application.ScreenUpdating = False
            
            '// ワークシートを追加。シート名はエクセルが命名
            Set wkSheet = ActiveWorkbook.Worksheets.Add(Count:=1)
            '// 結果表示
            headerRows = pfDrawHeader(wkSheet, rst)    '// ヘッダ行
            Call psDrawDataRows(wkSheet, rst, headerRows)  ', cmbGroup.Value)   '// データ行
            
            '// ページ設定
            Application.StatusBar = MSG_PAGE_SETUP
            Call gsPageSetup_Lines(wkSheet, headerRows)
            
            '// コメント設定
            Call Selection.NoteText("-- " & Format(Now, "yyyy/mm/dd hh:nn:ss") & vbCrLf & txtScript.Text)
            
            '// 警告表示
            If rst.Fields.Count > Columns.Count Then
              Call MsgBox(MSG_TOO_MANY_COLS, vbOKOnly, APP_TITLE)
            End If
            
            '// 書式の設定
            '//フォント
            wkSheet.Cells.Font.Name = APP_FONT
            wkSheet.Cells.Font.Size = APP_FONT_SIZE
            
            Call wkSheet.Cells(1, 1).Select
        Else
            Call MsgBox(MSG_NO_RESULT, vbOKOnly, APP_TITLE)
        End If
    Else    '// DMLの場合
        Call MsgBox(gADO.DmlRows & MSG_ROWS_PROCESSED, vbOKOnly, APP_TITLE)
    End If
    
  
    '// 後処理
    If rst.State = adStateOpen Then
        Call rst.Close
    End If
    
    Set rst = Nothing
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Call Me.Hide
    Exit Sub
  
ErrorHandler:
    Call gsShowErrorMsgDlg("frmGetRecord.psExecSearch", Err, gADO)
    Application.StatusBar = False
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// メソッド：    列ヘッダ描画
'// 説明：        列ヘッダを描画します。
'// 引数：        wkSheet: ワークシート
'//               rst: レコードセット
'// 戻り値：      ヘッダ行数
'// ////////////////////////////////////////////////////////////////////////////
Private Function pfDrawHeader(wkSheet As Worksheet, rst As Object) As Integer
On Error GoTo ErrorHandler
    Dim idx       As Integer
    Dim colStr    As String
    Dim strFormat As String
  
    '// ヘッダ描画行数（戻り値）を設定
    Select Case cmdHeader.Value
        Case 0
            pfDrawHeader = 1
        Case 1
            pfDrawHeader = 3
        Case 2
            pfDrawHeader = 0
    End Select
  
    '// ヘッダ行の項目
    For idx = 1 To IIf(rst.Fields.Count > Columns.Count, Columns.Count, rst.Fields.Count)
        With rst.Fields(idx - 1)
            '// 書式設定 //////////
            Select Case CLng(.Type)
                '// 2:adSmallInt, 3:adInteger, 4:adSingle, 5:adDouble, 6:adCurrency, 16:adTinyInt, 17:adUnsignedTinyInt, 18:adUnsignedSmallInt, 19:adUnsignedInt, 20: adBigInt, 21:adUnsignedBigInt, 131:adNumeric, 139:adVarNumeric
                Case 2, 3, 4, 5, 6, 16, 17, 18, 19, 20, 21, 131, 139
                    strFormat = BLANK
                Case 133, 135                     '// adDBDate, adDBTimeStamp
                    strFormat = cmbDateFormat.List(cmbDateFormat.ListIndex, 1)
                Case 134                          '// 134:adDBTime
                    strFormat = "hh:mm:ss"
                Case Else
                    strFormat = "@"
            End Select
            Call wkSheet.Columns(idx).Select
            Selection.NumberFormatLocal = strFormat
            
            '// 名称設定 //////////
            If cmdHeader.Value <> 2 Then
                wkSheet.Cells(1, idx).NumberFormatLocal = "@"
                wkSheet.Cells(1, idx).Value = .Name
            End If
            
            '// 定義の出力（型・桁数）//////////
            If cmdHeader.Value = 1 Then
                Select Case CLng(.Type)
                    Case 129, 130                     '// adChar, adWChar
                        wkSheet.Cells(2, idx).Value = "CHAR(" & .DefinedSize & ")"
                    Case 200, 202                     '//adVarChar, adVarWChar
                        wkSheet.Cells(2, idx).Value = "VARCHAR(" & .DefinedSize & ")"
                    Case 2, 18                        '// 2:adSmallInt, 18:adUnsignedSmallInt
                        wkSheet.Cells(2, idx).Value = "SMALLINT"
                    Case 3, 19                        '// 3:adInteger, 19:adUnsignedInt
                        wkSheet.Cells(2, idx).Value = "INTEGER"
                    Case 16, 17                       '// 16:adTinyInt 17:adUnsignedTinyInt
                        wkSheet.Cells(2, idx).Value = "TINYINT"
                    Case 20, 21                       '// 20:adBigInt, 21:adUnsignedBigInt
                        wkSheet.Cells(2, idx).Value = "BIGINT"
                    Case 4                            '// 4:adSingle
                        wkSheet.Cells(2, idx).Value = "SINGLE"
                    Case 5                            '// 5:adDouble
                        wkSheet.Cells(2, idx).Value = "DOUBLE"
                    Case 6                            '// 6:adCurrency
                        wkSheet.Cells(2, idx).Value = "CURRENCY"
                    Case 131, 139                     '// 131:adNumeric, 139:adVarNumeric
                        If .Precision = 0 Then
                            wkSheet.Cells(2, idx).Value = "NUMERIC"
                        ElseIf .NumericScale >= 0 Then
                            wkSheet.Cells(2, idx).Value = "NUMERIC(" & .Precision & "," & .NumericScale & ")"
                        Else
                            wkSheet.Cells(2, idx).Value = "NUMERIC(" & .Precision & ")"
                        End If
                    Case 133                          '// 133:adDBDate
                        wkSheet.Cells(2, idx).Value = "DATE"
                    Case 134                          '// 134:adDBTime
                        wkSheet.Cells(2, idx).Value = "TIME"
                    Case 135                          '// adDBTimeStamp
                        wkSheet.Cells(2, idx).Value = "TIMESTAMP"
                    Case 203  '// lob
                        wkSheet.Cells(2, idx).Value = "CLOB"
                    Case Else
                        wkSheet.Cells(2, idx).Value = BLANK
                End Select
                wkSheet.Cells(3, idx).Value = "-"
            End If
        End With
    Next
    '// 枠線の設定
    Call wkSheet.Range(Cells(1, pfDrawHeader + 1), Cells(1, wkSheet.UsedRange.Columns.Count)).Select
    Call gsDrawLine_Header
    
    '// 枠の固定を設定
    If pfDrawHeader > 0 Then
        Call wkSheet.Activate
        Call wkSheet.Rows(pfDrawHeader + 1).Select
        ActiveWindow.FreezePanes = True
    End If
    Exit Function

ErrorHandler:
    Call gsShowErrorMsgDlg("frmGetRecord.pfDrawHeader", Err)
End Function


'// ////////////////////////////////////////////////////////////////////////////
'// メソッド：    帳票描画
'// 説明：        各行の値を描画します。
'// 引数：        wksheet: ワークシート
'//               rst: レコードセット
'//               headerRows: ヘッダ行数
'//               groupIdx: グループ化する列数(V2で廃止）
'// 戻り値：      なし
'// ////////////////////////////////////////////////////////////////////////////
Private Sub psDrawDataRows(wkSheet As Worksheet, rst As Object, headerRows As Integer)  ', groupIdx As Integer)
On Error GoTo ErrorHandler
    Dim idxRow          As Long
    Dim idxCol          As Integer
    Dim cntCol          As Integer
    Dim varResult       As Variant    '// 結果保持配列（列,行）※redimの仕様対応のため、行と列を通常と反対に持つので注意
    
    idxRow = 0
  
    Do While Not rst.EOF
        '// Variant配列整備
        If idxRow = 0 Then
            cntCol = rst.Fields.Count
            ReDim varResult(cntCol - 1, 1)
            
        Else
            ReDim Preserve varResult(cntCol - 1, idxRow + 1)
        End If
        idxRow = idxRow + 1
        
        '// データを配列（列, 行）に格納
        For idxCol = 0 To IIf(cntCol > Columns.Count - 1, Columns.Count - 1, cntCol - 1)
            varResult(idxCol, idxRow - 1) = IIf(IsNull(rst.Fields(idxCol).Value), BLANK, rst.Fields(idxCol).Value)
'            If (idxCol > groupIdx) Then
'                '// 最小レベル以降の場合、値を描画
'                varResult(idxResult, idxCol) = IIf(IsNull(rst.Fields(idxCol - 1).Value), BLANK, rst.Fields(idxCol - 1).Value)
'            ElseIf (aryLastVal(idxCol - 1) = BLANK) Or (aryLastVal(idxCol - 1) <> rst.Fields(idxCol - 1).Value) Then
'                '// 直前の値が異なる場合 (含 直前の値がブランクの場合)
'                '// 配下のレベルの直前の値をクリア
'                For aryIdx = groupIdx To idxCol Step -1
'                    aryLastVal(aryIdx - 1) = BLANK
'                Next
'                varResult(idxResult, idxCol) = IIf(IsNull(rst.Fields(idxCol - 1).Value), BLANK, rst.Fields(idxCol - 1).Value)
'                aryLastVal(idxCol - 1) = IIf(IsNull(rst.Fields(idxCol - 1).Value), BLANK, rst.Fields(idxCol - 1).Value)
'            End If
        Next
        Call rst.MoveNext
    Loop
    
    '// Variantの内容を行列入れ替えてシートに張り付け
    wkSheet.Range(wkSheet.Cells(headerRows + 1, 1), wkSheet.Cells(idxRow + headerRows, cntCol)).Value = WorksheetFunction.Transpose(varResult)
    
    '// 罫線を描画
    Call wkSheet.UsedRange.Select
    Call gsDrawLine_Data
    
    Exit Sub

ErrorHandler:
    Call gsShowErrorMsgDlg("frmGetRecord.psDrawDataRows", Err)
End Sub

'// ////////////////////////////////////////////////////////////////////////////
'// END.
'// ////////////////////////////////////////////////////////////////////////////
