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
'// Windows API 関連の宣言
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
Public Declare PtrSafe Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As LongPtr, ByVal hwnd As LongPtr, ByVal Msg As LongPtr, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr


Private Const WM_NOTIFY = &H4E           '//0x004E
Private Const MCN_SELECT = -746


'// 日付時刻（MonthView からMCM_GETCURSEL指定で日付を取得する際に使用）
Private Type SYSTEMTIME
    wYear           As Integer
    wMonth          As Integer
    wDayOfWeek      As Integer
    wDay            As Integer
    wHour           As Integer
    wMinute         As Integer
    wSecond         As Integer
    wMilliseconds   As Integer
End Type

Private Type NMHDR
    hwndFrom        As LongPtr
    idFrom          As LongPtr
    code            As Long  'Integer
End Type

Private Type tagNMSELCHANGE
    hdr             As NMHDR
    stSelStart      As SYSTEMTIME
    stSelEnd        As SYSTEMTIME
End Type


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
Public Function WindowProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As LongPtr
On Error GoTo ErrorHandler
    Dim tagNMHDR    As NMHDR
    Dim prm         As tagNMSELCHANGE
    
    '// ToDo:　ブックが無いときの処理
    If uMsg = WM_NOTIFY Then
        Call CopyMemory(tagNMHDR, ByVal lParam, Len(tagNMHDR))
        If tagNMHDR.hwndFrom = hMonthView And tagNMHDR.code = MCN_SELECT Then
            Call CopyMemory(prm, ByVal lParam, Len(prm))
            ActiveCell.Value = CDate(prm.stSelStart.wYear & "/" & prm.stSelStart.wMonth & "/" & prm.stSelStart.wDay)
        End If
    End If
    '// 処理の成否にかかわらず必ずデフォルトのウィンドウプロシージャに引き渡す
ErrorHandler:
    WindowProc = CallWindowProc(defaultProcAddress, hwnd, uMsg, wParam, lParam)
End Function
