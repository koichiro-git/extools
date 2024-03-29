Attribute VB_Name = "mdlConvJapaneseNumber"
'// ////////////////////////////////////////////////////////////////////////////
'// プロジェクト   : 拡張ツール 追加パック
'// タイトル       : 日本語表記数値変換
'// モジュール     : mdlConvJapaneseNumber
'// 説明           : 日本語表記数値（「5百万」など）を数値に変換する。
'//                : 小数点（「5.5億」など）は現バージョンでは未対応。将来対応予定
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
    Dim cnt         As Integer  '// エラーカウント
    
    '// 事前チェック（アクティブシート保護、選択タイプ＝セル）
    If Not gfPreCheck(protectCont:=True, selType:=TYPE_RANGE) Then
        Exit Sub
    End If
    
    Call gsSuppressAppEvents
    
    cnt = 0
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
        ElseIf ActiveCell.Text <> BLANK Then    '// 変換に失敗し、セル値が空白でない場合にはエラー扱いとしてカウント
            cnt = cnt + 1
        End If
    End If
    
    Call gsResumeAppEvents
    
    If cnt > 0 Then
        Call MsgBox(MSG_ERR & "(" & cnt & ")", vbOKOnly, APP_TITLE)
    End If
    
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
Private Function pfConvJapaneseToNumber_sub(targetStr As String) As String
On Error GoTo ErrorHandler
    Dim i           As Integer  '// 文字を抽出する際のインデクス
    Dim c           As String   '// 抽出された一文字を格納するバッファ
    Dim dig         As Double   '// 現在の位
    Dim lastKanji   As String   '// 感じが現れた場合、その最後の漢字を保持（万、億、兆のいずれか）
    Dim result      As Double   '// 変換の経過を保持する
    
    If targetStr = BLANK Then
        pfConvJapaneseToNumber_sub = BLANK
        Exit Function
    End If
    
    dig = 1
    i = Len(targetStr)
    Do While i > 0
        c = Mid(targetStr, i, 1)
        
        If IsNumeric(c) Then
            result = result + dig * Int(c)
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
                
                Case "-", "−", "△", "▲", ","
                    '// マイナス記号,カンマは無視
                Case Else
                    '// それ以外の文字があった場合はエラーとしてブランクを戻す
                    pfConvJapaneseToNumber_sub = BLANK
                    Exit Function
            End Select
        End If
        
        i = i - 1
    Loop
    '// マイナス対応
    If InStr(1, "-▲△", Left(targetStr, 1)) > 0 Then
        result = result * -1
    End If
    pfConvJapaneseToNumber_sub = result
    Exit Function
    
ErrorHandler:
    pfConvJapaneseToNumber_sub = BLANK
    Call gsResumeAppEvents
    Call gsShowErrorMsgDlg_VBA("mdlConvJapaneseNumber.pfConvJapaneseToNumber_sub", Err)
End Function


'// ////////////////////////////////////////////////////////////////////////////
'// END
'// ////////////////////////////////////////////////////////////////////////////

