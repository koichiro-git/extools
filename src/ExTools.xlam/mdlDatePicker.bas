Attribute VB_Name = "mdlDatePicker"
Option Explicit
Option Base 0
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long


Private Const WM_NOTIFY = &H4E           '//0x004E
Private Const MCN_SELECT = -746


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

Private Type NMHDR
    hwndFrom    As Long
    idFrom      As Long
    code        As Long  'Integer
End Type

Private Type tagNMSELCHANGE
    hdr         As NMHDR
    stSelStart  As SYSTEMTIME
    stSelEnd    As SYSTEMTIME
End Type


Public defaultProcAddress As LongPtr
Public hMonthView      As LongPtr  '// MonthViewのウィンドウハンドル





Public Function WindowProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim tagNMHDR As NMHDR
    Dim prm As tagNMSELCHANGE
'    Dim x, y As Long
'    Dim tempdate  As SYSTEMTIME
    
    
    '// https://stackoverflow.com/questions/66578251/c-winapi-month-calendar-control
    '// https://gist.github.com/baoo777/1759063b5e90cc33a157592ab9b6adae   '// かなりいいかも
    Select Case uMsg
        Case WM_NOTIFY
            Call CopyMemory(tagNMHDR, ByVal lParam, Len(tagNMHDR))
            If tagNMHDR.hwndFrom = hMonthView Then
                If tagNMHDR.code = MCN_SELECT Then
                    Call CopyMemory(prm, ByVal lParam, Len(prm))
'                    Debug.Print (prm.stSelStart.wYear & prm.stSelStart.wMonth & prm.stSelStart.wDay)
                    ActiveCell.Value = CDate(prm.stSelStart.wYear & "/" & prm.stSelStart.wMonth & "/" & prm.stSelStart.wDay)
                End If
            End If
    End Select
    
    WindowProc = CallWindowProc(defaultProcAddress, hwnd, uMsg, wParam, lParam)

End Function
