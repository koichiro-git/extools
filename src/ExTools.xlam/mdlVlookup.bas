Attribute VB_Name = "mdlVlookup"
'// ////////////////////////////////////////////////////////////////////////////
'// プロジェクト   : 拡張ツール
'// タイトル       : Vlookup貼り付け機能
'// モジュール     : mdlVlookup
'// 説明           : VLOOKUPのマスタ定義の記憶、およびVLOOKUP関数の貼り付けを行う
'// ////////////////////////////////////////////////////////////////////////////
'// Copyright (c) by Koichiro.
'// ////////////////////////////////////////////////////////////////////////////
Option Explicit
Option Base 0

'// ////////////////////////////////////////////////////////////////////////////
'// プライベート変数
Private pVLookupMaster                  As String               '// VLookUpコピー機能でマスタ表範囲を格納する
Private pVLookupMasterIndex             As String               '// VLookUpコピー機能でマスタ表範囲の表示インデクスを格納する」


'// ////////////////////////////////////////////////////////////////////////////
'// メソッド：   リボンボタンコールバック管理(フォームなし)
'// 説明：       リボンからのコールバックをつかさどる
'//              押されたコントロールのIDを基に処理を呼び出す。
'// 引数：       control 対象コントロール
'// ////////////////////////////////////////////////////////////////////////////
Public Sub ribbonCallback_Vlookup(control As IRibbonControl)
    Select Case control.ID
        Case "VLookupCopy"                  '// VLookup
            Call psVLookupCopy
        Case "VLookupPaste"
            Call psVLookupPaste
    End Select
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// メソッド：   VLookupのマスタ領域としてコピー
'// 説明：       選択領域を表す文字列を内部変数に格納する
'// ////////////////////////////////////////////////////////////////////////////
Private Sub psVLookupCopy()
    '// 事前チェック
    If gfPreCheck(selType:=TYPE_RANGE, selAreas:=1) = False Then
        Exit Sub
    End If
    
    '// 1列のみの選択はエラー
    If Selection.Columns.Count = 1 Then
        Call MsgBox(MSG_VLOOKUP_MASTER_2COLS, vbOKOnly, APP_TITLE)
        Exit Sub
    End If
    
    pVLookupMaster = Selection.Worksheet.Name & "!" & Selection.Address(True, True)   '// 列と行を絶対参照
    pVLookupMasterIndex = Selection.Columns.Count
End Sub
            
            
'// ////////////////////////////////////////////////////////////////////////////
'// メソッド：   VLookup関数を張り付け
'// 説明：       VLookupCopyで格納された選択領域を張り付け位置のセルに出力する
'// ////////////////////////////////////////////////////////////////////////////
Private Sub psVLookupPaste()
On Error GoTo ErrorHandler
    Dim searchColIdx    As Long     '// VLookup関数の「検索」にあたるセルの列
    Dim targetColIdx    As Long     '// Vlookpu関数を出力するセルの列
    Dim bffRange        As String   '// 選択範囲のアドレス文字列を保持
    
    '// 事前チェック
    If gfPreCheck(protectCont:=True, selType:=TYPE_RANGE, selAreas:=1) = False Then
        Exit Sub
    End If
    
    '// マスタ列が選択されていない場合はエラー
    If pVLookupMaster = BLANK Then
        Call MsgBox(MSG_VLOOKUP_NO_MASTER, vbOKOnly, APP_TITLE)
        Exit Sub
    End If
    
    '// 1列のみの選択はエラー
    If Selection.Columns.Count = 1 Then
        Call MsgBox(MSG_VLOOKUP_SET_2COLS, vbOKOnly, APP_TITLE)
        Exit Sub
    End If
    
    '// 同一シート内での貼り付けの場合はマスタ表の範囲との重複をチェック
    If Selection.Worksheet.Name = Range(pVLookupMaster).Worksheet.Name Then
        If Not Application.Intersect(Selection, Range(Range(pVLookupMaster).Address)) Is Nothing Then
            Call MsgBox(MSG_VLOOKUP_SEL_DUPLICATED, vbOKOnly, APP_TITLE)
            Exit Sub
        End If
    End If
    
    '// 選択範囲のうち、カレントセルがVLOOKUPの「検索」列に該当
    searchColIdx = ActiveCell.Column
    '// 実際にVLOOKUP関数を埋め込むセルは、選択範囲の最後側。
    targetColIdx = IIf(Selection.Column = ActiveCell.Column, Selection.Column + Selection.Columns.Count - 1, Selection.Column)
    
    '// 最初のセルを出力
    Cells(Selection.Row, targetColIdx).Value = "=VLOOKUP(" & ActiveCell.Address(False, False) & "," & pVLookupMaster & "," & Str(pVLookupMasterIndex) & ",FALSE)"
    '// 下方向に数式のみコピー
    bffRange = Selection.Address(False, False)
    Call Cells(Selection.Row, targetColIdx).Copy
    Call Range(Cells(Selection.Row, targetColIdx), Cells(Selection.Row + Selection.Rows.Count - 1, targetColIdx)).PasteSpecial(Paste:=xlPasteFormulas, Operation:=xlNone, SkipBlanks:=False, Transpose:=False)
    
    '// 後処理
    Application.CutCopyMode = False
    Exit Sub
    
ErrorHandler:
    Call gsShowErrorMsgDlg("psVLookupPaste", Err)
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// END
'// ////////////////////////////////////////////////////////////////////////////

