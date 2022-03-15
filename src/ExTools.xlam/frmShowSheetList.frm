VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmShowSheetList 
   Caption         =   "シート一覧出力"
   ClientHeight    =   3375
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4305
   OleObjectBlob   =   "frmShowSheetList.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmShowSheetList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'// ////////////////////////////////////////////////////////////////////////////
'// プロジェクト   : 拡張ツール
'// タイトル       : シート一覧出力フォーム
'// モジュール     : frmShowSheetList
'// 説明           : シート一覧を出力する
'// ////////////////////////////////////////////////////////////////////////////
'// Copyright (c) by Koichiro.
'// ////////////////////////////////////////////////////////////////////////////
Option Explicit
Option Base 0


'// ////////////////////////////////////////////////////////////////////////////
'// プライベート定数
Private Const pDEF_ROWS  As Integer = 5
Private Const pDEF_COLS  As Integer = 5
Private Const pMAX_ROWS  As Integer = 10  '// 行数コンボボックスの最大値
Private Const pMAX_COLS  As Integer = 60  '// 列数コンボボックスの最大値


'// //////////////////////////////////////////////////////////////////
'// イベント： フォーム アクティブ時
Private Sub UserForm_Activate()
    '// ブックが開かれていない場合は終了
    If Workbooks.Count = 0 Then
        Call MsgBox(MSG_NO_BOOK, vbOKOnly, APP_TITLE)
        Call Me.Hide
        Exit Sub
    End If
    
    If ActiveWorkbook.MultiUserEditing Or (cmbOutput.Value = "0") Then
        ckbHyperLink.Value = False
        ckbHyperLink.Enabled = False
    Else
        ckbHyperLink.Enabled = True
    End If
End Sub


'// //////////////////////////////////////////////////////////////////
'// イベント： フォーム 初期化時
Private Sub UserForm_Initialize()
    Dim idx   As Integer
    
    Call gsSetCombo(cmbOutput, CMB_SSL_OUTPUT, 0)
    
    '// 行・列コンボの設定
    With cmbRows
        Call .Clear
        For idx = 0 To pMAX_ROWS
            Call .AddItem(CStr(idx))
            .List(idx, 1) = CStr(idx)
        Next
        .ListIndex = pDEF_ROWS
    End With
    
    With cmbCols
        Call .Clear
        For idx = 0 To pMAX_COLS
            Call .AddItem(CStr(idx))
            .List(idx, 1) = CStr(idx)
        Next
        .ListIndex = pDEF_COLS
    End With
    
    '// キャプション設定
    frmShowSheetList.Caption = LBL_SSL_FORM
    cmdExecute.Caption = LBL_COM_EXEC
    cmdClose.Caption = LBL_COM_CLOSE
    fraOption.Caption = LBL_SSL_OPTIONS
    ckbHyperLink.Caption = LBL_COM_HYPERLINK
    lblTarget.Caption = LBL_SSL_TARGET
    lblRows.Caption = LBL_SSL_ROWS
    lblCols.Caption = LBL_SSL_COLS
End Sub


'// //////////////////////////////////////////////////////////////////
'// イベント： 実行ボタン クリック時
Private Sub cmdExecute_Click()
    '// ブックが保護されている場合で、同一ブックに結果シートを追加しようとした場合はエラー
    If ActiveWorkbook.ProtectStructure And (cmbOutput.Value <> "0") Then
        Call MsgBox(MSG_BOOK_PROTECTED, vbOKOnly, APP_TITLE)
    Else
        '// 処理実施
        Call psShowSheetList(ActiveWorkbook, cmbRows.Value, cmbCols.Value)
    End If
    Call Me.Hide
End Sub


'// //////////////////////////////////////////////////////////////////
'// イベント： 閉じるボタン クリック時
Private Sub cmdClose_Click()
    Call Me.Hide
End Sub


'// //////////////////////////////////////////////////////////////////
'// イベント： 出力先コンボ 変更時
Private Sub cmbOutput_Change()
    '// ブックが開かれていない場合はなにもせず終了
    If Workbooks.Count = 0 Then
        Exit Sub
    End If

    If cmbOutput.Value = "0" Then
        ckbHyperLink.Value = False
        ckbHyperLink.Enabled = False
    ElseIf Not ActiveWorkbook.MultiUserEditing Then
        ckbHyperLink.Enabled = True
    End If
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// メソッド：   シート一覧出力
'// 説明：       シート一覧を出力する
'//              実行ボタンクリックイベントから呼び出される。
'// 引数：       wkBook: 対象ブック
'//              maxRow: 出力行数
'//              maxCol: 出力列数
'// ////////////////////////////////////////////////////////////////////////////
Private Sub psShowSheetList(wkBook As Workbook, maxRow As Integer, maxCol As Integer)
    Dim resultSheet As Worksheet  '// 結果出力先のシート
    Dim sheetObj    As Object     '// worksheet または chart オブジェクトを格納
    Dim idx         As Integer    '// 結果シートのカラム位置補正
    Dim idxRow      As Integer
    Dim idxCol      As Integer
    Dim statGauge   As cStatusGauge
    
    '// ブックが開かれていない場合は終了
    If Workbooks.Count = 0 Then
        Call MsgBox(MSG_NO_BOOK, vbOKOnly, APP_TITLE)
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    Set statGauge = New cStatusGauge
    statGauge.MaxVal = wkBook.Sheets.Count * maxRow
  
  '// 出力先の設定
    Select Case cmbOutput.Value
        Case "0"
            Call Workbooks.Add
            Set resultSheet = ActiveWorkbook.ActiveSheet
        Case "1"
            Set resultSheet = ActiveWorkbook.Worksheets.Add(Before:=ActiveWorkbook.Worksheets(1))
        Case "2"
            Set resultSheet = ActiveWorkbook.Worksheets.Add(After:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count))
    End Select
  
    '// B列はシート名称のため、書式を文字列（@）に設定
    resultSheet.Columns("B").NumberFormat = "@"
    
    '// ヘッダの設定(セルインデクスの表示は共通関数を使用しない。最初の2列分（"シート番号;シート名称"）のみをHDR_SSLで設定)
    Call gsDrawResultHeader(resultSheet, HDR_SSL, 1)
    '// 3列目以降のヘッダは「A1,B1,C1...」をコンボボックスでの指定列分設定
    idx = 3
    For idxRow = 1 To maxRow
        For idxCol = 1 To maxCol
            resultSheet.Cells(1, idx).Value = gfGetColIndexString(idxCol) & CStr(idxRow)
            idx = idx + 1
        Next
    Next
  
    '// 一覧（データ部）の出力
    For Each sheetObj In wkBook.Sheets
        resultSheet.Cells(sheetObj.Index + 1, 1).Value = sheetObj.Index
        resultSheet.Cells(sheetObj.Index + 1, 2).Value = sheetObj.Name
        
        If sheetObj.Type = xlWorksheet Then '// ワークシートのみ、内容の表示とリンクの設定
            '// リンクの設定
            If ckbHyperLink.Value And (sheetObj.Visible = xlSheetVisible) Then
                Call Cells(resultSheet.Index + 1, 2).Hyperlinks.Add(Anchor:=Cells(sheetObj.Index + 1, 2), Address:="", SubAddress:="'" & sheetObj.Name & "'!A1")
            End If
          
            '// シート設定値の出力
            If maxRow * maxCol > 0 Then
                idx = 3
                For idxRow = 1 To maxRow
                    For idxCol = 1 To maxCol
                        resultSheet.Cells(sheetObj.Index + 1, idx).NumberFormat = sheetObj.Cells(idxRow, idxCol).NumberFormat
                        resultSheet.Cells(sheetObj.Index + 1, idx).Value = sheetObj.Cells(idxRow, idxCol).Value
                        idx = idx + 1
                    Next
                    Call statGauge.addValue(1)
                Next
            End If
        End If
    Next
  
    Call gsPageSetup_Lines(resultSheet, 1)
  
    Set statGauge = Nothing
    '// 別ブックに出力した際には、閉じるときに保存を求めない
    ActiveWorkbook.Saved = (cmbOutput.Value = "0")
    Application.ScreenUpdating = True
End Sub

'// ////////////////////////////////////////////////////////////////////////////
'// END
'// ////////////////////////////////////////////////////////////////////////////
