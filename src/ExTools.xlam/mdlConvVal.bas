Attribute VB_Name = "mdlConvVal"
'// ////////////////////////////////////////////////////////////////////////////
'// プロジェクト   : 拡張ツール
'// タイトル       : 文字種の変換
'// モジュール     : mdlConvVal
'// 説明           : セル文字列の変換、トリム機能
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
Public Sub ribbonCallback_ConvVal(control As IRibbonControl)
    Select Case control.ID
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
    End Select
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
    
    Call gsSuppressAppEvents
    
    If Selection.Count > 1 Then
        For Each tCell In Selection.SpecialCells(xlCellTypeConstants, xlNumbers + xlTextValues)
            Call psConvValue_sub(tCell, funcFlag)
            
            '// キー割込
            If GetAsyncKeyState(27) <> 0 Then
                Application.StatusBar = False
                Exit For
            End If
        Next
    Else
        Call psConvValue_sub(ActiveCell, funcFlag)
    End If
    
    Call gsResumeAppEvents
    Exit Sub
ErrorHandler:
    Call gsResumeAppEvents
    If Err.Number = 1004 Then  '// 範囲選択が正しくない場合
        Call MsgBox(MSG_INVALID_RANGE, vbOKOnly, APP_TITLE)
    Else
        Call gsShowErrorMsgDlg("psConvValue", Err)
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
'// END
'// ////////////////////////////////////////////////////////////////////////////
