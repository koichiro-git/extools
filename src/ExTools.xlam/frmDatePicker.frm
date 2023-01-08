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

Private Const CALENDAR_SEP_WIDTH    As Double = 6  '// カレンダー2つ（2か月）分の間隔4.5pt + 予備


'// //////////////////////////////////////////////////////////////////
'// イベント： フォーム 初期化時
Private Sub UserForm_Initialize()
    Call psSetupDatePicker
End Sub


Private Sub psSetupDatePicker()
On Error GoTo ErrorHandler
    Dim icce                As tagINITCOMMONCONTROLSEX
    Dim rc                  As RECT
    Dim hWnd_Sub            As LongPtr
    Dim hWnd                As LongPtr  '// UserFormのウィンドウハンドル
    Dim lResult             As LongPtr
    Dim calendarWidth       As Long
    
    ' コモンコントロール初期化
    icce.dwICC = ICC_DATE_CLASSES
    icce.dwSize = Len(icce)
    lResult = InitCommonControlsEx(icce)
    If lResult = 0 Then Call Err.Raise(Number:=513, Description:="日付ピッカー画面を生成できません")
    
    ' ユーザーフォームのHWNDの取得
    hWnd = FindWindow(THUNDER_FRAME, Me.Caption)
    If hWnd = 0 Then Call Err.Raise(Number:=513, Description:="日付ピッカー画面を生成できません")

    ' MonthView配置用ハンドルの取得
    hWnd_Sub = FindWindowEx(hWnd, 0, vbNullString, vbNullString)
    
    '// MonthViewウィンドウ生成(サイズゼロで生成　https://learn.microsoft.com/ja-jp/windows/win32/controls/mcm-getminreqrect)
    hMonthView = CreateWindowEx(0, MONTHCAL_CLASS, vbNullString, (WS_VISIBLE Or WS_CHILD), 0, 0, 0, 0, hWnd_Sub, 0, 0, vbNullString) '//hWnd_Sub, 0, lnghInstance, vbNullString)
    '// MonthView用ウィンドウのリサイズ
    lResult = SendMessage(hMonthView, MCM_GETMINREQRECT, 0, rc)
    calendarWidth = (rc.Right - rc.Left) * 2 + CALENDAR_SEP_WIDTH
    Call MoveWindow(hMonthView, 0, 0, calendarWidth, rc.Bottom - rc.Top, 1&)

    defaultProcAddress = SetWindowLongPtr(hWnd_Sub, GWL_WNDPROC, AddressOf WindowProc)

    '// 画面補正 //////////
    Call SetWindowLongPtr(hWnd, GWL_EXSTYLE, GetWindowLongPtr(hWnd, GWL_EXSTYLE) Or WS_EX_TOOLWINDOW)   '// UserFormをToolWindowスタイルに変更
    
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
