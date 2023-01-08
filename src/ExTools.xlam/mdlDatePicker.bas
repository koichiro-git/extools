Attribute VB_Name = "mdlDatePicker"
'// ////////////////////////////////////////////////////////////////////////////
'// プロジェクト   : 拡張ツール
'// タイトル       : 日付ピッカー用標準モジュール
'// モジュール     : mdlDatePicker
'// 説明           : 日付ピッカーフォームで使用するウィンドウプロシージャを記載
'// ////////////////////////////////////////////////////////////////////////////
'// Copyright (c) by Koichiro.
'// ////////////////////////////////////////////////////////////////////////////
Option Explicit
Option Base 0

'// ////////////////////////////////////////////////////////////////////////////
'// 変数
Public defaultProcAddress   As LongPtr
Public hMonthView           As LongPtr  '// MonthViewのウィンドウハンドル


'// ////////////////////////////////////////////////////////////////////////////
'// メソッド：   日付ピッカー用ウィンドウプロシージャ
'// 説明：       日付選択時の処理
'// 引数：       hwnd:   ウィンドウハンドル
'//              uMsg:   メッセージ
'//              wParam: 追加のメッセージ固有情報
'//              lParam: 追加のメッセージ固有情報
'// ////////////////////////////////////////////////////////////////////////////
Public Function WindowProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As LongPtr
On Error GoTo ErrorHandler
    Dim tagNMHDR    As NMHDR
    Dim prm         As tagNMSELCHANGE
    
    If uMsg = WM_NOTIFY Then
        Call CopyMemory(tagNMHDR, ByVal lParam, Len(tagNMHDR))
        If tagNMHDR.hwndFrom = hMonthView And tagNMHDR.code = MCN_SELECT Then
            Call CopyMemory(prm, ByVal lParam, Len(prm))
            ActiveCell.Value = CDate(prm.stSelStart.wYear & "/" & prm.stSelStart.wMonth & "/" & prm.stSelStart.wDay)
        End If
    End If
    '// アクティブセルが無い場合などもエラーはすべて無視されるため条件判定はしない。
    '// 処理の成否にかかわらず必ずデフォルトのウィンドウプロシージャに引き渡すため、Exitは無し
    
ErrorHandler:
    WindowProc = CallWindowProc(defaultProcAddress, hWnd, uMsg, wParam, lParam)
End Function
