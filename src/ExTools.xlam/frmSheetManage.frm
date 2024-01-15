VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSheetManage 
   Caption         =   "シート操作"
   ClientHeight    =   5055
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6360
   OleObjectBlob   =   "frmSheetManage.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmSheetManage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'// ////////////////////////////////////////////////////////////////////////////
'// プロジェクト   : 拡張ツール
'// タイトル       : シート設定フォーム
'// モジュール     : frmSheetManage
'// 説明           : シートの初期化を行う
'// ////////////////////////////////////////////////////////////////////////////
'// Copyright (c) by Koichiro.
'// ////////////////////////////////////////////////////////////////////////////
Option Explicit
Option Base 0


'// //////////////////////////////////////////////////////////////////
'// イベント： フォーム アクティブ時
Private Sub UserForm_Activate()
    '// 事前チェック（シート有無）
    If Not gfPreCheck() Then
        Call Me.Hide
        Exit Sub
    End If
End Sub


'// //////////////////////////////////////////////////////////////////
'// イベント： フォーム初期化時
Private Sub UserForm_Initialize()
    '// コンボボックス設定
    Call gsSetCombo(cmbTarget, CMB_SMG_TARGET, 0)
    Call gsSetCombo(cmbView, CMB_SMG_VIEW, 0)
    Call gsSetCombo(cmbZoom, CMB_SMG_ZOOM, 0)
    Call gsSetCombo(cmbFilter, CMB_SMG_FILTER, 0)
    
    '// キャプション設定
    frmSheetManage.Caption = LBL_SMG_FORM
    ckbScroll.Caption = LBL_SMG_SCROLL
    ckbFontColor.Caption = LBL_SMG_FONT_COLOR
    ckbLink.Caption = LBL_SMG_HYPERLINK
    ckbComment.Caption = LBL_SMG_COMMENT
    ckbHeader.Caption = LBL_SMG_HEAD_FOOT
    ckbMargin.Caption = LBL_SMG_MARGIN
    ckdbPageBreak.Caption = LBL_SMG_PAGEBREAK
    fraPrintOpt.Caption = LBL_SMG_PRINT_OPT
    optPrintNone.Caption = LBL_SMG_PRINT_NONE
    optPrintNoZoom.Caption = LBL_SMG_PRINT_100
    optPrintVert.Caption = LBL_SMG_PRINT_HRZ
    optPrint1Page.Caption = LBL_SMG_PRINT_1_PAGE
    lblTarget.Caption = LBL_SMG_TARGET
    lblView.Caption = LBL_SMG_VIEW
    lblZoom.Caption = LBL_SMG_ZOOM
    lblFilter.Caption = LBL_SMG_AUTOFILTER
    cmdExecute.Caption = LBL_COM_EXEC
    cmdClose.Caption = LBL_COM_CLOSE
    cmdSelectAll.Caption = LBL_COM_CHECK_ALL
    cmdClear.Caption = LBL_COM_UNCHECK
End Sub


'// //////////////////////////////////////////////////////////////////
'// イベント： すべて選択ボタン クリック時
Private Sub cmdSelectAll_Click()
    Call setCheckBoxes(True)
End Sub


'// //////////////////////////////////////////////////////////////////
'// イベント： 選択解除ボタン クリック時
Private Sub cmdClear_Click()
    Call setCheckBoxes(False)
End Sub


'// //////////////////////////////////////////////////////////////////
'// イベント： 実行ボタン クリック時
Private Sub cmdExecute_Click()
On Error GoTo ErrorHandler
    Dim wkBook      As Workbook
    Dim wkSheet     As Worksheet
    Dim FilePath    As String
    Dim FileName    As String
    Dim compFiles   As String   '// 複数ファイル更新時、完了したファイル名を保持
  
    '// Zoom値のチェック
    If IsNull(cmbZoom.Value) Then
      If IsNumeric(cmbZoom.Text) Then
        If CInt(cmbZoom.Text) < 10 Or CInt(cmbZoom.Text) > 400 Then
          Call MsgBox(MSG_VAL_10_400, vbOKOnly, APP_TITLE)
          Exit Sub
        End If
      Else
          Call MsgBox(MSG_VAL_10_400, vbOKOnly, APP_TITLE)
        Exit Sub
      End If
    End If
  
    '// ブックの保護確認
    Select Case cmbTarget.Value
        Case 0 '// 現在のシート
            If ActiveSheet.ProtectContents Then
                Call MsgBox(MSG_SHEET_PROTECTED, vbOKOnly, APP_TITLE)
                Exit Sub
            End If
        Case 1 '// 現在のワークブック
            For Each wkSheet In Worksheets
                If wkSheet.ProtectContents Then
                    Call MsgBox(MSG_SHEETS_PROTECTED, vbOKOnly, APP_TITLE)
                    Exit Sub
                End If
            Next
    End Select
  
    Call gsSuppressAppEvents
    
    '// 処理対象コンボの値によって{シート | ブック | ディレクトリ単位}に実行
    Select Case cmbTarget.Value
        Case 0    '// 現在のシート
            Call psSetUpSheetProperty(ActiveSheet)
        Case 1    '// 現在のブック
            For Each wkSheet In Worksheets
                Call psSetUpSheetProperty(wkSheet)
            Next
            Call ActiveWorkbook.Sheets(1).Activate
        Case 2    '// ディレクトリ単位
            '// 実行前設定
            If Not gfShowSelectFolder(0, FilePath) Then
                Exit Sub
            End If
            
            FileName = Dir(FilePath & "\*.xls")
            Do While FileName <> BLANK
                Call Workbooks.Open(FilePath & "\" & FileName, ReadOnly:=False)
                For Each wkSheet In ActiveWorkbook.Worksheets
                    Call psSetUpSheetProperty(wkSheet)
                Next
                Call ActiveWorkbook.Sheets(1).Activate
                compFiles = compFiles & ActiveWorkbook.Name & Chr(10)
                Call ActiveWorkbook.Close(SaveChanges:=True)
                FileName = Dir
            Loop
            
            '// xlsx形式（Excel2007以上）への対応
            FileName = Dir(FilePath & "\*.xlsx")
            Do While FileName <> BLANK
                Call Workbooks.Open(FilePath & "\" & FileName, ReadOnly:=False)
                For Each wkSheet In ActiveWorkbook.Worksheets
                    Call psSetUpSheetProperty(wkSheet)
                Next
                Call ActiveWorkbook.Sheets(1).Activate
                compFiles = compFiles & Chr(10) & ActiveWorkbook.Name
                Call ActiveWorkbook.Close(SaveChanges:=True)
                
                FileName = Dir
            Loop
    End Select
    
    Call gsResumeAppEvents
    Call MsgBox(MSG_FINISHED, vbOKOnly, APP_TITLE)
    
    Call Me.Hide
    Exit Sub
  
ErrorHandler:
    Call gsResumeAppEvents
    Call gsShowErrorMsgDlg("frmSheetManage.cmdExecute_Click [" & ActiveWorkbook.FullName & "!" & ActiveSheet.Name & "]", Err, Nothing)
    Call MsgBox(MSG_COMPLETED_FILES & Chr(10) & compFiles, vbOKOnly, APP_TITLE)
End Sub


'// //////////////////////////////////////////////////////////////////
'// イベント： 閉じるボタン クリック時
Private Sub cmdClose_Click()
    Call Me.Hide
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// メソッド：   チェックボックス設定
'// 説明：       チェックボックスの設定を引数の真偽値に設定する。
'// 引数：       setVal: 設定値
'// 戻り値：     なし
'// ////////////////////////////////////////////////////////////////////////////
Private Sub setCheckBoxes(setVal As Boolean)
    ckbScroll.Value = setVal
    ckbFontColor.Value = setVal
    ckbLink.Value = setVal
    ckbComment.Value = setVal
    ckbHeader.Value = setVal
    ckbMargin.Value = setVal
    ckdbPageBreak.Value = setVal
    cmbView.Value = IIf(setVal, cmbView.Value, 0)
  
    '// 以下の項目については「選択解除(setVal=false)」の場合のみ補正
    If setVal = False Then
        cmbZoom.Value = 0
        optPrintNone.Value = True
        cmbFilter.Value = 0
    End If
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// メソッド：   シート設定
'// 説明：       引数のシートの整形処理を行う
'//              実行ボタンクリックイベントから呼び出される
'// 引数：       wkSheet: 対象ワークシート
'// ////////////////////////////////////////////////////////////////////////////
Private Sub psSetUpSheetProperty(wkSheet As Worksheet)
    Call wkSheet.Activate
    Call wkSheet.Cells(1, 1).Select  '// グラフなどがアクティブな場合のエラーを回避するため、A1を選択状態にする
    Application.StatusBar = "Setting up: [" & wkSheet.Parent.Name & "!" & wkSheet.Name & "]"
    
    '// ビューを設定 ※スクロール設定よりも先にビューを変更する必要あり（FreezePanes設定でエラーとなるため）
    Select Case cmbView.Value
        Case 1
            ActiveWindow.View = xlNormalView
        Case 2
            ActiveWindow.View = xlPageBreakPreview
    End Select
    
    '// スクロールの初期化
    If ckbScroll.Value Then
        ActiveWindow.FreezePanes = False
        ActiveWindow.Split = False
        ActiveWindow.ScrollColumn = 1
        ActiveWindow.ScrollRow = 1
    End If
    
    '// ズームを初期化
    If IsNull(cmbZoom.Value) Or cmbZoom.Value <> 0 Then
        ActiveWindow.Zoom = cmbZoom.Text
    End If
    
    '// オートフィルタの設定 「0:指定無し」は無視
    Select Case cmbFilter.Value
        Case 1 '// フィルタ解除
            If wkSheet.AutoFilterMode Then
                Call wkSheet.Cells.AutoFilter
            End If
        Case 2 '// 全て表示
            If WorksheetFunction.CountA(ActiveSheet.UsedRange) > 1 Then
                Call wkSheet.ShowAllData
            End If
        Case 3 '// １行目でフィルタ
            If Not wkSheet.AutoFilterMode And WorksheetFunction.CountA(ActiveSheet.UsedRange) > 1 Then
                Call wkSheet.Cells.AutoFilter
            End If
    End Select
    
    '// ハイパーリンクを削除
    If ckbLink.Value Then
        Call wkSheet.Hyperlinks.Delete
    End If
    
    '// 印刷の拡大/縮小
    With wkSheet.PageSetup
        If optPrintNoZoom.Value Then
            .Zoom = 100
        ElseIf optPrintVert.Value Then
            .Zoom = False
            .FitToPagesWide = 1
            .FitToPagesTall = False
        ElseIf optPrint1Page.Value Then
            .Zoom = False
            .FitToPagesWide = 1
            .FitToPagesTall = 1
        End If
    End With
    
    '// ヘッダとフッタの設定
    If ckbHeader.Value Then
        Call mdlCommon.gsPageSetup_Header(wkSheet)
    End If
    
    '// マージンの設定
    If ckbMargin.Value Then
        Call mdlCommon.gsPageSetup_Margin(wkSheet)
    End If
    
    '// 改ページと印刷範囲を解除
    If ckdbPageBreak.Value Then
        Call ActiveSheet.ResetAllPageBreaks
        wkSheet.PageSetup.PrintArea = ""
    End If
    
    '// フォント色の初期化
    If ckbFontColor.Value Then
        wkSheet.Cells.Font.ColorIndex = xlAutomatic
    End If
    
    '// コメントを削除
    If ckbComment.Value Then
        Call wkSheet.Cells.ClearComments
    End If
    
    Call wkSheet.Cells(1, 1).Select
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// END
'// ////////////////////////////////////////////////////////////////////////////
