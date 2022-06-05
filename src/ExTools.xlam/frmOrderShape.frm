VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmOrderShape 
   Caption         =   "シェイプの配置"
   ClientHeight    =   3330
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4305
   OleObjectBlob   =   "frmOrderShape.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmOrderShape"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'// ////////////////////////////////////////////////////////////////////////////
'// プロジェクト   : 拡張ツール
'// タイトル       : シェイプの整列フォーム
'// モジュール     : frmOrderShape
'// 説明           : 選択されたシェイプをセルにあわせて整列させる
'// ////////////////////////////////////////////////////////////////////////////
'// Copyright (c) by Koichiro.
'// ////////////////////////////////////////////////////////////////////////////
Option Explicit
Option Base 0


'// //////////////////////////////////////////////////////////////////
'// イベント： フォーム 初期化時
Private Sub UserForm_Initialize()
    Dim idx   As Integer
    
    '// コンボボックス設定
    Call gsSetCombo(cmbHeight, CMB_ORD_HEIGHT, 0)
    Call gsSetCombo(cmbWidth, CMB_ORD_WIDTH, 0)
    
    With cmbMargin
        Call .Clear
        For idx = 0 To 8
            Call .AddItem(CStr(idx))
            .List(idx, 1) = CStr(idx) & " pt"
        Next
        .ListIndex = 1
    End With
    
    Call ckbDetail_Change   '// チェックボックス
    
    '// キャプション設定
    frmOrderShape.Caption = LBL_ORD_FORM
    cmdExecute.Caption = LBL_COM_EXEC
    cmdClose.Caption = LBL_COM_CLOSE
    ckbDetail.Caption = LBL_ORD_OPTIONS
    lblMargin.Caption = LBL_ORD_MARGIN
    lblHeight.Caption = LBL_ORD_HEIGHT
    lblWidth.Caption = LBL_ORD_WIDTH
End Sub


'// //////////////////////////////////////////////////////////////////
'// イベント： 閉じるボタン クリック時
Private Sub cmdClose_Click()
    Call Me.Hide
End Sub


'// //////////////////////////////////////////////////////////////////
'// イベント： 詳細設定チェックボックス 変更時
Private Sub ckbDetail_Change()
    '// セルにフィットさせるときのみ有効
    cmbHeight.Enabled = ckbDetail.Value
    cmbWidth.Enabled = ckbDetail.Value
End Sub


'// //////////////////////////////////////////////////////////////////
'// イベント： 実行ボタン クリック時
Private Sub cmdExecute_Click()
On Error GoTo ErrorHandler
    Dim idx   As Integer
    
    For idx = 1 To ActiveWindow.Selection.ShapeRange.Count  '// shaperangeの開始インデックスは１から
        Call psSetShapePos(ActiveWindow.Selection.ShapeRange(idx), cmbMargin.Value)
    Next
    Exit Sub
  
ErrorHandler:
    Call MsgBox(MSG_SHAPE_NOT_SELECTED, vbOKOnly, APP_TITLE)
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// メソッド：   シェイプ位置・サイズ設定
'// 説明：
'// 引数：       targetShape: 対象シェイプオブジェクト
'//              ptMargin: マージン
'// ////////////////////////////////////////////////////////////////////////////
Private Sub psSetShapePos(targetShape As Shape, ptMargin As Integer)
    Dim basePos(4) As Long
    
    basePos(0) = targetShape.TopLeftCell.Top + ptMargin '// 上端
    basePos(1) = targetShape.TopLeftCell.Left + ptMargin '// 左端
    basePos(2) = targetShape.BottomRightCell.Top + targetShape.BottomRightCell.Height - ptMargin '// 下端
    basePos(3) = targetShape.BottomRightCell.Left + targetShape.BottomRightCell.Width - ptMargin '// 右端
    
    If targetShape.Type <> msoLine Then   '// 直線シェイプ以外を対象とする
        '// 上下端設定
        If Not ckbDetail.Value Or (cmbHeight.Value = 0) Or (cmbHeight.Value = 1) Then
            targetShape.Top = basePos(0)
            If cmbHeight.Value = 0 Then
                targetShape.Height = basePos(2) - basePos(0)
            End If
        ElseIf cmbHeight.Value = 2 Then
            targetShape.Top = basePos(2) - targetShape.Height
        End If
        
        '// 左右端設定
        If Not ckbDetail.Value Or (cmbWidth.Value = 0) Or (cmbWidth.Value = 1) Then
            targetShape.Left = basePos(1)
            If cmbWidth.Value = 0 Then
                targetShape.Width = basePos(3) - basePos(1)
            End If
        ElseIf cmbWidth.Value = 2 Then
            targetShape.Left = basePos(3) - targetShape.Width
        End If
    End If
End Sub

'// ////////////////////////////////////////////////////////////////////////////
'// END
'// ////////////////////////////////////////////////////////////////////////////
