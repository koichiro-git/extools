Attribute VB_Name = "mdlWindowsAPI"
'// ////////////////////////////////////////////////////////////////////////////
'// プロジェクト   : 拡張ツール
'// タイトル       : Windows API宣言
'// モジュール     : mdlWindowsAPI
'// 説明           : Windows API 関連の宣言。Win32API_PtrSafe.txtベース
'// ////////////////////////////////////////////////////////////////////////////
'// Copyright (c) by Koichiro.
'// ////////////////////////////////////////////////////////////////////////////
Option Explicit
Option Base 0

'// ////////////////////////////////////////////////////////////////////////////
'// 定数

'// Window Styles
Public Const WS_VISIBLE = &H10000000
Public Const WS_CHILD = &H40000000
'Private Const WS_BORDER = &H800000
Public Const WS_EX_TOOLWINDOW = &H80

'// Window field offsets for GetWindowLong() and GetWindowWord()
Public Const GWL_WNDPROC = (-4)
Public Const GWL_EXSTYLE = (-20)

'// MonthView SendMessage用メッセージ定義
Public Const MCM_FIRST = &H1000
Public Const MCM_GETCURSEL = (MCM_FIRST + 1)           '// 選択された日付を取得
Public Const MCM_SETCURSEL = (MCM_FIRST + 2)
Public Const MCM_GETMINREQRECT = (MCM_FIRST + 9)       '// MonthViewのサイズを取得

'// MonthView関連
Public Const MONTHCAL_CLASS = "SysMonthCal32"
Public Const ICC_DATE_CLASSES = &H100
Public Const WM_NOTIFY = &H4E           '//0x004E
Public Const MCN_SELECT = -746

'// フォルダ選択ダイアログ関連
Public Const BIF_RETURNONLYFSDIRS = &H1
Public Const MAX_PATH = 260


'// ////////////////////////////////////////////////////////////////////////////
'// タイプ

'// InitCommonControlsEx用
Public Type tagINITCOMMONCONTROLSEX
    dwSize          As Long
    dwICC           As Long
End Type

'// 日付時刻（MonthView からMCM_GETCURSEL指定で日付を取得する際に使用）
Public Type SYSTEMTIME
    wYear           As Integer
    wMonth          As Integer
    wDayOfWeek      As Integer
    wDay            As Integer
    wHour           As Integer
    wMinute         As Integer
    wSecond         As Integer
    wMilliseconds   As Integer
End Type

'// RECT（ウィンドウサイズ設定時に使用。単位＝ピクセル）
Public Type RECT
    Left            As Long
    Top             As Long
    Right           As Long
    Bottom          As Long
End Type

'// 座標 （マウス座標）
Public Type POINTAPI
    x               As Long
    y               As Long
End Type

'// フォルダ選択ダイアログ関連
Public Type BROWSEINFO
    hwndOwner       As LongPtr
    pidlRoot        As Long
    pszDisplayName  As String
    lpszTitle       As String
    ulFlags         As Long
    lpfn            As LongPtr
    lParam          As LongPtr
    iImage          As Long
End Type

'// ウインドウメッセージ管理関連
Public Type NMHDR
    hwndFrom        As LongPtr
    idFrom          As LongPtr
    code            As Long
End Type

Public Type tagNMSELCHANGE
    hdr             As NMHDR
    stSelStart      As SYSTEMTIME
    stSelEnd        As SYSTEMTIME
End Type


'// ////////////////////////////////////////////////////////////////////////////
'// 関数

'// ウインドウ生成
Public Declare PtrSafe Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As LongPtr, ByVal hMenu As LongPtr, ByVal hInstance As LongPtr, lpParam As Any) As LongPtr
'// ウインドウ座標設定
Public Declare PtrSafe Function MoveWindow Lib "user32" (ByVal hWnd As LongPtr, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
'// マウス座標取得
Public Declare PtrSafe Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
'// ウィンドウコントロールの操作
Public Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, lParam As Any) As LongPtr
'// ウィンドウハンドル取得
Public Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
Public Declare PtrSafe Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As LongPtr, ByVal hWnd2 As LongPtr, ByVal lpsz1 As String, ByVal lpsz2 As String) As LongPtr

'// コモンコントロール初期化
Public Declare PtrSafe Function InitCommonControlsEx Lib "ComCtl32" (LPINITCOMMONCONTROLSEX As Any) As Long

'// ウィンドウスタイル補正
#If Win64 Then
    Public Declare PtrSafe Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongPtrA" (ByVal hWnd As LongPtr, ByVal nIndex As Long) As LongPtr
    Public Declare PtrSafe Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongPtrA" (ByVal hWnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
#Else
    Public Declare PtrSafe Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As LongPtr, ByVal nIndex As Long) As LongPtr
    Public Declare PtrSafe Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
#End If

'// メモリコピー（キャスト代替）
Public Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
'// ウィンドウ プロシージャ呼び出し
Public Declare PtrSafe Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As LongPtr, ByVal hWnd As LongPtr, ByVal Msg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr

'// フォルダ選択
Public Declare PtrSafe Function apiSHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolder" (lpBrowseInfo As BROWSEINFO) As LongPtr
'// パス取得
Public Declare PtrSafe Function apiSHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDList" (ByVal piDL As LongPtr, ByVal strPath As String) As LongPtr
'//キー割り込み
Public Declare PtrSafe Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
