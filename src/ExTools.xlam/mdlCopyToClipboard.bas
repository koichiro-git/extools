Attribute VB_Name = "mdlCopyToClipboard"
'// ////////////////////////////////////////////////////////////////////////////
'// プロジェクト   : 拡張ツール
'// タイトル       : クリップボードへコピー機能
'// モジュール     : mdlCopyToClipboard
'// 説明           : 選択範囲を固定長、マークダウン形式、または画像形式でクリップボードにコピーする
'// ////////////////////////////////////////////////////////////////////////////
'// Copyright (c) by Koichiro.
'// ////////////////////////////////////////////////////////////////////////////
Option Explicit
Option Base 0

'// ////////////////////////////////////////////////////////////////////////////
'// コンパイルスイッチ（"EXCEL" / "POWERPOINT"）
#Const OFFICE_APP = "EXCEL"

'// ////////////////////////////////////////////////////////////////////////////
'// アプリケーション定数
Private Const MAX_COL_LEN             As Integer = 80                                               '// クリップボードにコピーする際の列最大長


'// ////////////////////////////////////////////////////////////////////////////
'// メソッド：   リボンボタンコールバック管理(フォームなし)
'// 説明：       リボンからのコールバックをつかさどる
'//              押されたコントロールのIDを基に処理を呼び出す。
'// 引数：       control 対象コントロール
'// ////////////////////////////////////////////////////////////////////////////
Public Sub ribbonCallback_Copy2CB(control As IRibbonControl)
    Select Case control.ID
#If OFFICE_APP = "EXCEL" Then
        Case "Copy2Clipboard"               '// 固定長コピー
            Call psCopyToClipboard
        Case "Copy2CBMarkdown"              '// マークダウン形式でコピー
            Call psCopyToCB_Markdown
        Case "Copy2CBImage"                 '// 画像としてコピー
            Call psCopyToCB_Image
#End If
        Case "Copy2CBShapeText"             '// シェイプのテキストをコピー
            Call psCopyShapeText
    End Select
End Sub


#If OFFICE_APP = "EXCEL" Then
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
    
    '// 事前チェック（選択タイプ＝シェイプ）
    If Not gfPreCheck(selType:=TYPE_RANGE) Then
        Exit Sub
    End If
    
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
    Exit Sub
    
ErrorHandler:
    Call gsShowErrorMsgDlg("psCopyToClipboard", Err, Nothing)
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
    
    '// 事前チェック（選択タイプ＝セル）
    If Not gfPreCheck(selType:=TYPE_RANGE) Then
        Exit Sub
    End If
  
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
    Call gsShowErrorMsgDlg("psCopyToCB_Markdown", Err, Nothing)
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// メソッド：   クリップボードへイメージ形式でコピー
'// 説明：       選択範囲をイメージ形式でクリップボードに格納する。
'// 引数：       なし
'// 戻り値：     なし
'// ////////////////////////////////////////////////////////////////////////////
Public Sub psCopyToCB_Image()
On Error GoTo ErrorHandler
        
    '// 事前チェック（選択タイプ＝セル）
    If Not gfPreCheck(selType:=TYPE_RANGE) Then
        Exit Sub
    End If
    
    '// 選択範囲が単一であることの確認
    If Selection.Areas.Count > 1 Then
        Call MsgBox(MSG_TOO_MANY_RANGE, vbOKOnly, APP_TITLE)
        Exit Sub
    End If
    '// コピー
    Call Selection.CopyPicture(xlScreen, xlBitmap)
        
    Exit Sub
ErrorHandler:
    Call gsShowErrorMsgDlg("psCopyToCB_Image", Err, Nothing)
End Sub
#End If


'// ////////////////////////////////////////////////////////////////////////////
'// 説明：       Win10 から DataObject.PutInClipboard が効かなくなったため、回避策としてテキストボックスを経由してコピー
'// 引数：       コピー対象文字列
'// ////////////////////////////////////////////////////////////////////////////
Private Sub psSetClip(bffText As String)
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
Private Function pfIsPercentage(bffText As String) As Boolean
    If bffText = BLANK Then
        pfIsPercentage = False
    ElseIf Right(bffText, 1) = "%" And IsNumeric(Left(bffText, Len(bffText) - 1)) Then
        pfIsPercentage = True
    Else
        pfIsPercentage = False
    End If
End Function


'// ////////////////////////////////////////////////////////////////////////////
'// メソッド：   シェイプ内のテキストをコピー
'// 説明：       ネストしたグループ内もすべてコピーする。実体は_subに実装
'// ////////////////////////////////////////////////////////////////////////////
Private Sub psCopyShapeText()
On Error GoTo ErrorHandler
    Dim idx         As Integer
    Dim sh          As Shape
    Dim bff         As String
    
    '// 事前チェック（選択タイプ＝シェイプ）
    If Not gfPreCheck(selType:=TYPE_SHAPE) Then
        Exit Sub
    End If
    
#If OFFICE_APP = "EXCEL" Then
    For idx = 1 To Selection.ShapeRange.Count  '// shaperangeの開始インデックスは１から
#ElseIf OFFICE_APP = "POWERPOINT" Then
    For idx = 1 To ActiveWindow.Selection.ShapeRange.Count  '// shaperangeの開始インデックスは１から
#End If
        bff = bff & pfCopyShapeText_sub(ActiveWindow.Selection.ShapeRange(idx))
    Next
    
    Call psSetClip(bff)
    Exit Sub
    
ErrorHandler:
    Call gsShowErrorMsgDlg("psCopyShapeText", Err, Nothing, idx, sh.Name, bff)
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// メソッド：   シェイプ内のテキストをコピー
'// 説明：       シェイプのテキストコピー実装部
'// ////////////////////////////////////////////////////////////////////////////
Private Function pfCopyShapeText_sub(targetShape As Shape) As String
    Dim sh      As Shape
    Dim rslt    As String
    Dim bff     As String
    
    bff = BLANK
    '// グループは再帰処理
    If targetShape.Type = msoGroup Then
        For Each sh In targetShape.GroupItems
            bff = bff & pfCopyShapeText_sub(sh)
        Next
    End If
    
    bff = bff & gfClean(gfGetShapeText(targetShape))
'    bff = WorksheetFunction.Clean(gfGetShapeText(targetShape))
    If bff <> BLANK Then
        pfCopyShapeText_sub = rslt & Trim(Str(Int(targetShape.Left))) & "," + Trim(Str(Int(targetShape.Top))) & "," & bff & vbCrLf
    End If
End Function


'// ////////////////////////////////////////////////////////////////////////////
'// END
'// ////////////////////////////////////////////////////////////////////////////

