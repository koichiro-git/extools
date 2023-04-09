VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAdjustArch 
   Caption         =   "円弧の調整"
   ClientHeight    =   2280
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3135
   OleObjectBlob   =   "frmAdjustArch.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmAdjustArch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'// ////////////////////////////////////////////////////////////////////////////
'// プロジェクト   : 拡張ツール
'// タイトル       : 円弧の調整フォーム
'// モジュール     : frm
'// 説明           : 円弧オブジェクトの開始位置、終了位置を角度で指定する
'//                : 対象とするシェイプは Pie, BlockArc, CircularArrow, msoShapeArc
'// ////////////////////////////////////////////////////////////////////////////
'// Copyright (c) by Koichiro.
'// ////////////////////////////////////////////////////////////////////////////
Option Explicit
Option Base 0

Private Const ANGLE_ADJUST      As Integer = -90    '// 角度計算の開始位置補正値

Private angleStart              As Integer          '// 開始角度
Private angleEnd                As Integer          '// 終了角度


'// //////////////////////////////////////////////////////////////////
'// イベント： 終了角度テキストボックス KeyDown時
Private Sub txtEnd_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
On Error GoTo ErrorHandler
    Dim lastVal     As Integer
    
    lastVal = angleEnd
    '// Enter, Tab, テンキーEnerの場合、描画処理を行う
    If KeyCode = vbKeyReturn Or KeyCode = vbKeySeparator Or KeyCode = vbKeyTab Then
        angleEnd = Int(txtEnd.Value)
        Call adjustArch
    End If
    Exit Sub

ErrorHandler:
    Call gsResumeAppEvents
    If Err.Number = 13 Or Err.Number = 6 Then  '// 入力した数値が無効な場合
        Call MsgBox(MSG_INVALID_NUM, vbOKOnly, APP_TITLE)
    Else
        Call gsShowErrorMsgDlg("frmAdjustArch.txtEnd_KeyDown", Err)
    End If
    
    angleEnd = lastVal
    txtEnd.Value = lastVal
End Sub


'// //////////////////////////////////////////////////////////////////
'// イベント： 開始角度テキストボックス KeyDown時
Private Sub txtStart_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
On Error GoTo ErrorHandler
    Dim lastVal     As Integer
    
    lastVal = angleStart
    '// Enter, Tab, テンキーEnerの場合、描画処理を行う
    If KeyCode = vbKeyReturn Or KeyCode = vbKeySeparator Or KeyCode = vbKeyTab Then
        angleStart = Int(txtStart.Value)
        Call adjustArch
    End If
    Exit Sub

ErrorHandler:
    Call gsResumeAppEvents
    If Err.Number = 13 Or Err.Number = 6 Then  '// 入力した数値が無効な場合
        Call MsgBox(MSG_INVALID_NUM, vbOKOnly, APP_TITLE)
    Else
        Call gsShowErrorMsgDlg("mdlCommon.psConvValue", Err)
    End If
    
    angleStart = lastVal
    txtStart.Value = lastVal
End Sub


'// //////////////////////////////////////////////////////////////////
'// イベント： フォーム 初期化時
Private Sub UserForm_Initialize()
    '// キャプション設定
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
    Dim spacer  As Integer
    
    '// 事前準備 //////////
    '// 事前チェック（アクティブシート保護、選択タイプ＝シェイプ）
    If Not gfPreCheck(protectCont:=True, selType:=TYPE_SHAPE) Then
        Me.Hide
        Exit Sub
    End If
    
    For Each shp In Selection.ShapeRange
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
    Call gsResumeAppEvents
    Call gsShowErrorMsgDlg("frmAdjustArch.UserForm_Activate", Err)
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
    
    For Each shp In Selection.ShapeRange
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
    
    Call adjustArch
End Sub
'// SpinDown
Private Sub spnStart_SpinDown()
    If angleStart = Int(angleStart / 15) * 15 Then
        angleStart = Int(angleStart / 15) * 15 - 15
    Else
        angleStart = Int(angleStart / 15) * 15
    End If
    
    txtStart.Value = angleStart
    Call adjustArch
End Sub


'// //////////////////////////////////////////////////////////////////
'// イベント：スピンボタン(終了角度)
'// SpinUp
Private Sub spnEnd_SpinUp()
    angleEnd = Int(angleEnd / 15) * 15 + 15
    txtEnd.Value = angleEnd
    
    Call adjustArch
End Sub
'// SpinDown
Private Sub spnEnd_SpinDown()
    If angleEnd = Int(angleEnd / 15) * 15 Then
        angleEnd = Int(angleEnd / 15) * 15 - 15
    Else
        angleEnd = Int(angleEnd / 15) * 15
    End If
    
    txtEnd.Value = angleEnd
    Call adjustArch
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// メソッド：   円弧角度調整
'// 説明：       選択されている円弧オブジェクトの開始角度・終了角度を設定する
'// ////////////////////////////////////////////////////////////////////////////
Private Sub adjustArch()
    Dim shp     As Shape
    
    For Each shp In Selection.ShapeRange
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
