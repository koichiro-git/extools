Attribute VB_Name = "mdlHyperlink"
'// ////////////////////////////////////////////////////////////////////////////
'// プロジェクト   : 拡張ツール
'// タイトル       : ハイパーリンク
'// モジュール     : mdlHyperlink
'// 説明           : ハイパーリンクの設定および解除機能
'// ////////////////////////////////////////////////////////////////////////////
'// Copyright (c) by Koichiro.
'// ////////////////////////////////////////////////////////////////////////////
Option Explicit
Option Base 0


'// ////////////////////////////////////////////////////////////////////////////
'// メソッド：   リボンボタンコールバック管理(フォームなし)
'// 説明：       リボンからのコールバックをつかさどる
'//              押されたコントロールのIDを基に処理を呼び出す。
'// 引数：       control 対象コントロール
'// ////////////////////////////////////////////////////////////////////////////
Public Sub ribbonCallback_Hyperlink(control As IRibbonControl)
    Select Case control.ID
        Case "AddLink"                      '// リンクの追加
            Call psSetHyperLink
        Case "RemoveLink"                   '// リンクの削除
            Call psRemoveHyperLink
    End Select
End Sub


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
'// END
'// ////////////////////////////////////////////////////////////////////////////

