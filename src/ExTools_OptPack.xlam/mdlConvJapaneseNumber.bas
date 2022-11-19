Attribute VB_Name = "mdlConvJapaneseNumber"
'// ////////////////////////////////////////////////////////////////////////////
'// プロジェクト   : 拡張ツール 追加パック
'// タイトル       : 日本語表記数値変換
'// モジュール     : mdlConvJapaneseNumber
'// 説明           : システムの共通関数、起動時の設定などを管理
'//                  国内電話番号のみ対応。国番号（+81等）は対応していない
'// ////////////////////////////////////////////////////////////////////////////
'// Copyright (c) by Koichiro.
'// ////////////////////////////////////////////////////////////////////////////
Option Explicit
Option Base 0

Public Sub test()
    Dim Text As String
    Call pfConvJapaneseToNumber
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// メソッド：   リボンボタンコールバック管理
'// 説明：       リボンからのコールバックをつかさどる
'//              押されたコントロールのIDを基に処理を呼び出す。
'// 引数：       control 対象コントロール
'// ////////////////////////////////////////////////////////////////////////////
Public Sub ribbonCallback_JapaneseNum(control As IRibbonControl)
    Call pfConvJapaneseToNumber
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// メソッド：   日本語表記数値変換 主関数
'// 説明：       引数の日本語表記数値を数値に変換する（5万4千3百 → 54300）
'//              小数は非対応
'// ////////////////////////////////////////////////////////////////////////////
Private Sub pfConvJapaneseToNumber()
On Error GoTo ErrorHandler
    Dim tCell       As Range    '// 変換対象セル
    Dim bff         As String   '// 変換後文字列格納バッファ
    
    '// 事前チェック（アクティブシート保護、選択タイプ＝セル）
    If Not gfPreCheck(protectCont:=True, selType:=TYPE_RANGE) Then
        Exit Sub
    End If
    
    Call gsSuppressAppEvents
    
    If Selection.Count > 1 Then
        For Each tCell In Selection.SpecialCells(xlCellTypeConstants, xlNumbers + xlTextValues)  '//SELECTIONが空の場合はエラーハンドラでキャッチ
            bff = pfConvJapaneseToNumber_sub(tCell.Text)
            If bff <> BLANK Then    '// 変換ロジックからブランクが戻された場合は無視
                tCell.Value = bff
            End If
        Next
    Else
        bff = pfConvJapaneseToNumber_sub(ActiveCell.Text)
        If bff <> BLANK Then        '// 変換ロジックからブランクが戻された場合は無視
            ActiveCell.Value = bff
        End If
    End If
    
    Call gsResumeAppEvents
    Exit Sub
ErrorHandler:
    Call gsResumeAppEvents
    If Err.Number = 1004 Then  '// 範囲選択が正しくない場合
        Call MsgBox(MSG_INVALID_RANGE, vbOKOnly, APP_TITLE)
    Else
        Call gsShowErrorMsgDlg_VBA("mdlConvJapaneseNumber.pfConvJapaneseToNumber", Err)
    End If
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// メソッド：   日本語表記数値変換
'// 説明：       引数の日本語表記数値を数値に変換する（5万4千3百 → 54300）
'//              psFormatPhoneNumbers から呼び出される実処理
'// 引数：       日本語表記数値
'// ////////////////////////////////////////////////////////////////////////////
Private Function pfConvJapaneseToNumber_sub(targetStr As String) As Double
On Error GoTo ErrorHandler
'    Dim targetStr   As String
    Dim i           As Integer  '// 文字を抽出する際のインデクス
    Dim c           As String   '// 抽出された一文字を格納するバッファ
    Dim dig         As Double   '// 現在の位
    Dim lastKanji   As String   '// 感じが現れた場合、その最後の漢字を保持（万、億、兆のいずれか）
    Dim Result      As Double   '// 変換の経過を保持する

'    targetStr = "123兆４千5百億3千万5千"
'    targetStr = "2億3千万円"
'    targetStr = "230000000"
'    targetStr = "2"
'    targetStr = "-2百"
'    targetStr = "12兆400億325万5千4"
'    targetStr = "5万4千"
'    targetStr = "△12３兆4千00億3千2百5万5千4百"
    dig = 1
    i = Len(targetStr)
    Do While i > 0
        c = Mid(targetStr, i, 1)
        
        If IsNumeric(c) Then
            Result = Result + dig * Int(c)
            dig = dig * 10
        Else
            Select Case c
                Case "百"
                    Select Case lastKanji
                        Case ""
                            dig = 100
                        Case "万"
                            dig = 1000000
                        Case "億"
                            dig = 10000000000#
                        Case "兆"
                            dig = 100000000000000#
                    End Select
                Case "千"
                    Select Case lastKanji
                        Case ""
                            dig = 1000
                        Case "万"
                            dig = 10000000
                        Case "億"
                            dig = 100000000000#
                        Case "兆"
                            '// 100兆以上はエラー
                    End Select
                Case "万"
                    dig = 10000
                    lastKanji = "万"
                Case "億"
                    dig = 100000000
                    lastKanji = "億"
                Case "兆"
                    dig = 1000000000000#
                    lastKanji = "兆"
                '// これ以外の文字は無視
            End Select
        End If
        
        i = i - 1
    Loop
    '// マイナス対応
    If InStr(1, "-▲△", Left(targetStr, 1)) > 0 Then
        Result = Result * -1
    End If
    pfConvJapaneseToNumber_sub = Result
    Exit Function
    
ErrorHandler:
    Call gsResumeAppEvents
    Call gsShowErrorMsgDlg_VBA("mdlConvJapaneseNumber.pfConvJapaneseToNumber_sub", Err)
End Function


'// ////////////////////////////////////////////////////////////////////////////
'// END
'// ////////////////////////////////////////////////////////////////////////////

