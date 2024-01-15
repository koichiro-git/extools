VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAdjustArc 
   Caption         =   "円弧の調整"
   ClientHeight    =   2565
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3030
   OleObjectBlob   =   "frmAdjustArc.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmAdjustArc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'// ////////////////////////////////////////////////////////////////////////////
'// プロジェクト   : 拡張ツール
'// タイトル       : 円弧の調整フォーム
'// モジュール     : frmAdjustArc
'// 説明           : 円弧オブジェクトの開始位置、終了位置を角度で指定する
'//                : 対象とするシェイプは Pie, BlockArc, CircularArrow, msoShapeArc
'// ////////////////////////////////////////////////////////////////////////////
'// Copyright (c) by Koichiro.
'// ////////////////////////////////////////////////////////////////////////////
Option Explicit
Option Base 0

'// ////////////////////////////////////////////////////////////////////////////
'// コンパイルスイッチ（"EXCEL" / "POWERPOINT"）
'#Const OFFICE_APP = "EXCEL"

Private Const ANGLE_ADJUST      As Integer = -90    '// 角度計算の開始位置補正値

Private angleStart              As Integer          '// 開始角度
Private angleEnd                As Integer          '// 終了角度


'// //////////////////////////////////////////////////////////////////
'// イベント： フォーム 初期化時
Private Sub UserForm_Initialize()
    '// キャプション設定
    Me.Caption = LBL_ARC_FORM
    lblStart.Caption = LBL_ARC_START
    lblEnd.Caption = LBL_ARC_END
    cmdResetRotation.Caption = LBL_ARC_RESET_ROT
    cmdClose.Caption = LBL_COM_CLOSE
End Sub


'// //////////////////////////////////////////////////////////////////
'// イベント： フォーム アクティベート時
Private Sub UserForm_Activate()
On Error GoTo ErrorHandler
    Dim shp     As Shape
    
    '// 事前準備 //////////
    '// 事前チェック（アクティブシート保護、選択タイプ＝シェイプ）
    If Not gfPreCheck(protectCont:=True, selType:=TYPE_SHAPE) Then
        Me.Hide
        Exit Sub
    End If
    
    For Each shp In ActiveWindow.Selection.ShapeRange
        Select Case shp.AutoShapeType
            Case msoShapePie, msoShapeBlockArc, msoShapeArc
                angleStart = Int(shp.Adjustments.Item(1)) - ANGLE_ADJUST
                angleEnd = Int(shp.Adjustments.Item(2)) - ANGLE_ADJUST
                Exit For
            Case msoShapeCircularArrow
                angleStart = Int(shp.Adjustments.Item(4)) - ANGLE_ADJUST
                angleEnd = Int(shp.Adjustments.Item(3)) - ANGLE_ADJUST + shp.Adjustments.Item(2)
                Exit For
        End Select
    Next
    txtStart.Value = angleStart
    txtEnd.Value = angleEnd
    
    Exit Sub

ErrorHandler:
    Call gsShowErrorMsgDlg("frmAdjustArc.UserForm_Activate", Err, Nothing)
End Sub


'// //////////////////////////////////////////////////////////////////
'// イベント： 閉じるボタン クリック時
Private Sub cmdClose_Click()
    Call Me.Hide
End Sub


'// //////////////////////////////////////////////////////////////////
'// イベント： 傾きリセットボタン クリック時
Private Sub cmdResetRotation_Click()
    Dim shp As Shape
    
    For Each shp In ActiveWindow.Selection.ShapeRange
        Select Case shp.AutoShapeType
            Case msoShapePie, msoShapeBlockArc, msoShapeCircularArrow, msoShapeArc
                shp.Rotation = 0
        End Select
    Next
End Sub


'// //////////////////////////////////////////////////////////////////
'// イベント：スピンボタン(開始角度)
'// SpinUp
Private Sub spnStart_SpinUp()
    angleStart = Int(angleStart / 15) * 15 + 15
    txtStart.Value = angleStart
    
    Call adjustArc
End Sub
'// SpinDown
Private Sub spnStart_SpinDown()
    If angleStart = Int(angleStart / 15) * 15 Then
        angleStart = Int(angleStart / 15) * 15 - 15
    Else
        angleStart = Int(angleStart / 15) * 15
    End If
    
    txtStart.Value = angleStart
    Call adjustArc
End Sub


'// //////////////////////////////////////////////////////////////////
'// イベント：スピンボタン(終了角度)
'// SpinUp
Private Sub spnEnd_SpinUp()
    angleEnd = Int(angleEnd / 15) * 15 + 15
    txtEnd.Value = angleEnd
    
    Call adjustArc
End Sub
'// SpinDown
Private Sub spnEnd_SpinDown()
    If angleEnd = Int(angleEnd / 15) * 15 Then
        angleEnd = Int(angleEnd / 15) * 15 - 15
    Else
        angleEnd = Int(angleEnd / 15) * 15
    End If
    
    txtEnd.Value = angleEnd
    Call adjustArc
End Sub


'// //////////////////////////////////////////////////////////////////
'// イベント： 開始角度テキストボックス AfterUpdate時
Private Sub txtStart_AfterUpdate()
On Error GoTo ErrorHandler
    If IsNumeric(txtStart.Value) Then
        txtStart.Value = Int(txtStart.Value)
        angleStart = Int(txtStart.Value)
        Call adjustArc
    Else
        Call MsgBox(MSG_INVALID_NUM, vbOKOnly, APP_TITLE)
        txtStart.Value = ActiveWindow.Selection.ShapeRange(1).Adjustments.Item(1) - ANGLE_ADJUST
    End If
    Exit Sub
    
ErrorHandler:
    Call gsShowErrorMsgDlg("frmAdjustArc.txtStart_AfterUpdate", Err, Nothing)
End Sub


'// //////////////////////////////////////////////////////////////////
'// イベント： 終了角度テキストボックス AfterUpdate時
Private Sub txtEnd_AfterUpdate()
On Error GoTo ErrorHandler
    If IsNumeric(txtEnd.Value) Then
        txtEnd.Value = Int(txtEnd.Value)
        angleEnd = Int(txtEnd.Value)
        Call adjustArc
    Else
        Call MsgBox(MSG_INVALID_NUM, vbOKOnly, APP_TITLE)
        txtEnd.Value = ActiveWindow.Selection.ShapeRange(1).Adjustments.Item(2) - ANGLE_ADJUST
    End If
    Exit Sub
    
ErrorHandler:
    Call gsShowErrorMsgDlg("frmAdjustArc.txtEnd_AfterUpdate", Err, Nothing)
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// メソッド：   円弧角度調整
'// 説明：       選択されている円弧オブジェクトの開始角度・終了角度を設定する
'// ////////////////////////////////////////////////////////////////////////////
Private Sub adjustArc()
    Dim shp     As Shape
    
    For Each shp In ActiveWindow.Selection.ShapeRange
        Select Case shp.AutoShapeType
            Case msoShapePie, msoShapeBlockArc, msoShapeArc
                shp.Adjustments.Item(1) = angleStart + ANGLE_ADJUST
                shp.Adjustments.Item(2) = angleEnd + ANGLE_ADJUST
            Case msoShapeCircularArrow
                shp.Adjustments.Item(4) = angleStart + ANGLE_ADJUST
                shp.Adjustments.Item(3) = angleEnd + ANGLE_ADJUST - shp.Adjustments.Item(2)
        End Select
    Next
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// END
'// ////////////////////////////////////////////////////////////////////////////
