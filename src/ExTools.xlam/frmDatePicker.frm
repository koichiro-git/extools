VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDatePicker 
   Caption         =   "日付"
   ClientHeight    =   6060
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13890
   OleObjectBlob   =   "frmDatePicker.frx":0000
   ShowModal       =   0   'False
End
Attribute VB_Name = "frmDatePicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'// ////////////////////////////////////////////////////////////////////////////
'// プロジェクト   : 拡張ツール
'// タイトル       : 日付ピックカレンダー
'// モジュール     : frmDatePicker
'// 説明           : MonthViewを使用したDatePicker
'// ////////////////////////////////////////////////////////////////////////////
'// Copyright (c) by Koichiro.
'// ////////////////////////////////////////////////////////////////////////////
Option Explicit
Option Base 0

'// ////////////////////////////////////////////////////////////////////////////
'// アプリケーション定数

Private Const THUNDER_FRAME     As String = "ThunderDFrame" '// Excel VBAユーザーフォームのクラス名（Excel2000以降=ThunerDFrame / それ以前=ThunderXFrame）

'// 論理インチ当たりの画面のピクセル数（ポイント→ピクセル換算係数）
'// GetDeviceCaps が常に96を返すため､プログラムでの動的取得をやめ、定数とする（SetProcessDPIAwareは実装しない）
'// https://learn.microsoft.com/ja-jp/windows-hardware/manufacture/desktop/dpi-related-apis-and-registry-settings?view=windows-11
Private Const LOG_PIXELS        As Long = 96

Private Const CALENDAR_SEP_WIDTH    As Double = 6  ''''// カレンダー2つ（2か月）分の間隔4.5pt +　予備


'// ////////////////////////////////////////////////////////////////////////////
'// Windows API 関連の宣言

'// 定数
Private Const MONTHCAL_CLASS = "SysMonthCal32"
Private Const ICC_DATE_CLASSES = &H100          '// コモンコントロール用定数（日付と時刻の選択コントロール）

'// Window Styles
Private Const WS_VISIBLE = &H10000000
Private Const WS_CHILD = &H40000000
'Private Const WS_BORDER = &H800000
Private Const WS_EX_TOOLWINDOW = &H80

'// Window field offsets for GetWindowLong() and GetWindowWord()
Private Const GWL_WNDPROC = (-4)
Private Const GWL_EXSTYLE = (-20)

'// MonthView SendMessage用メッセージ定義
Private Const MCM_FIRST = &H1000
Private Const MCM_GETCURSEL = (MCM_FIRST + 1)           '// 選択された日付を取得
Private Const MCM_SETCURSEL = (MCM_FIRST + 2)
Private Const MCM_GETMINREQRECT = (MCM_FIRST + 9)       '// MonthViewのサイズを取得

'// タイプ
'// InitCommonControlsEx用
Private Type tagINITCOMMONCONTROLSEX
    dwSize          As Long
    dwICC           As Long
End Type

'// 日付時刻（MonthView からMCM_GETCURSEL指定で日付を取得する際に使用）
Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

'// RECT（ウィンドウサイズ設定時に使用。単位＝ピクセル）
Private Type RECT
    Left        As Long
    Top         As Long
    Right       As Long
    Bottom      As Long
End Type

'// 座標 （マウス座標）
Private Type POINTAPI
    x           As Long
    y           As Long
End Type

'// ウインドウ生成
Private Declare PtrSafe Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As LongPtr, ByVal hMenu As LongPtr, ByVal hInstance As LongPtr, lpParam As Any) As LongPtr
'// ウインドウ座標設定
Private Declare PtrSafe Function MoveWindow Lib "user32" (ByVal hWnd As LongPtr, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
'// マウス座標取得
Private Declare PtrSafe Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
'// ウィンドウコントロールの操作
Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, lParam As Any) As LongPtr
'// ウィンドウ存在判定
'Private Declare PtrSafe Function IsWindow Lib "user32" (ByVal hwnd As LongPtr) As Long
'// ウィンドウ廃棄（二重起動時）
'Private Declare PtrSafe Function DestroyWindow Lib "user32" (ByVal hwnd As LongPtr) As Long
'// ウィンドウハンドル取得
Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
Private Declare PtrSafe Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As LongPtr, ByVal hWnd2 As LongPtr, ByVal lpsz1 As String, ByVal lpsz2 As String) As LongPtr
'Private Declare PtrSafe Function GetActiveWindow Lib "user32" () As LongPtr


'// コモンコントロール初期化
Private Declare PtrSafe Function InitCommonControlsEx Lib "ComCtl32" (LPINITCOMMONCONTROLSEX As Any) As Long

'// ウィンドウスタイル補正
#If Win64 Then
    Private Declare PtrSafe Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongPtrA" (ByVal hWnd As LongPtr, ByVal nIndex As Long) As LongPtr
    Private Declare PtrSafe Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongPtrA" (ByVal hWnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
#Else
    Private Declare PtrSafe Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As LongPtr, ByVal nIndex As Long) As LongPtr
    Private Declare PtrSafe Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
#End If
    
'// ////////////////////////////////////////////////////////////////////////////
'// 変数
'Private hwndMonthView           As LongPtr  '// MonthViewのウィンドウハンドル


'// //////////////////////////////////////////////////////////////////
'// イベント： フォーム 初期化時
Private Sub UserForm_Initialize()
    Call psSetupDatePicker
End Sub


Private Sub psSetupDatePicker()
On Error GoTo ErrorHandler
    Dim icce                As tagINITCOMMONCONTROLSEX
    Dim rc                  As RECT
    Dim lnghWnd_Sub         As LongPtr
    Dim hwndForm            As LongPtr  '// UserFormのウィンドウハンドル
    Dim lResult             As LongPtr
    Dim calendarWidth       As Long
    
    ' コモンコントロール初期化
    icce.dwICC = ICC_DATE_CLASSES
    icce.dwSize = Len(icce)
    lResult = InitCommonControlsEx(icce)
    If lResult = 0 Then Call Err.Raise(Number:=513, Description:="日付ピッカー画面を生成できません")
    
    ' ユーザーフォームのHWNDの取得
    hwndForm = FindWindow(THUNDER_FRAME, Me.Caption)
    If hwndForm = 0 Then Call Err.Raise(Number:=513, Description:="日付ピッカー画面を生成できません")

    ' MonthView配置用ハンドルの取得
    lnghWnd_Sub = FindWindowEx(hwndForm, 0, vbNullString, vbNullString)
    
    '// MonthViewウィンドウ生成(サイズゼロで生成　https://learn.microsoft.com/ja-jp/windows/win32/controls/mcm-getminreqrect)
    hMonthView = CreateWindowEx(0, MONTHCAL_CLASS, vbNullString, (WS_VISIBLE Or WS_CHILD), 0, 0, 0, 0, lnghWnd_Sub, 0, 0, vbNullString) '//lnghWnd_Sub, 0, lnghInstance, vbNullString)
'    hMonthView = hwndMonthView
    '// MonthView用ウィンドウのリサイズ
    lResult = SendMessage(hMonthView, MCM_GETMINREQRECT, 0, rc)
    calendarWidth = (rc.Right - rc.Left) * 2 + CALENDAR_SEP_WIDTH
    Call MoveWindow(hMonthView, 0, 0, calendarWidth, rc.Bottom - rc.Top, 1&)

    defaultProcAddress = SetWindowLongPtr(lnghWnd_Sub, GWL_WNDPROC, AddressOf WindowProc)

    '// 画面補正 //////////
    Call SetWindowLongPtr(hwndForm, GWL_EXSTYLE, GetWindowLongPtr(hwndForm, GWL_EXSTYLE) Or WS_EX_TOOLWINDOW)   '// UserFormをToolWindowスタイルに変更
    
    '// フォームサイズ補正(px→pt変換)。横幅はカレンダー幅＋フォーム枠幅。縦高はカレンダー高＋フォーム枠幅
    Me.Width = calendarWidth * 72 / LOG_PIXELS + (Me.Width - Me.InsideWidth)
    Me.Height = rc.Bottom * 72 / LOG_PIXELS + (Me.Height - Me.InsideHeight)
    
    '// フォーム位置補正（マウス座標へ）
    Call MoveFormToMouse
    Exit Sub

ErrorHandler:
    Call gsResumeAppEvents
    Call gsShowErrorMsgDlg("frmDatePicker.psSetupDatePicker", Err)
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// メソッド：    フォーム位置補正
'// 説明：        フォームをマウス位置に移動させる
'// 引数：        なし
'// ////////////////////////////////////////////////////////////////////////////
Private Sub MoveFormToMouse()
    Dim mousePos As POINTAPI
    
    Call GetCursorPos(mousePos)
    Me.Left = 72 / LOG_PIXELS * mousePos.x
    Me.Top = 72 / LOG_PIXELS * mousePos.y
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// END.
'// ////////////////////////////////////////////////////////////////////////////
