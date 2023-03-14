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
'// //////////////////////////////////////////////////////////////////
'// 対象とするシェイプは Pie, BlockArc, CircularArrow

Option Explicit
Option Base 0

Private Const ANGLE_ADJUST      As Integer = -90  '// 角度計算の開始位置補正値

Private angleStart              As Integer      '// 開始角度
Private angleEnd                As Integer      '// 終了角度


Private Sub txtEnd_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Or KeyCode = vbKeySeparator Or KeyCode = vbKeyTab Then
        angleEnd = txtEnd.Value
        Call adjustArch
    End If
End Sub


Private Sub txtStart_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Or KeyCode = vbKeySeparator Or KeyCode = vbKeyTab Then
        angleStart = txtStart.Value
        Call adjustArch
    End If
End Sub

'// //////////////////////////////////////////////////////////////////
'// イベント： フォーム 初期化時
Private Sub UserForm_Initialize()
    cmdClose.Caption = LBL_COM_CLOSE
End Sub


'// //////////////////////////////////////////////////////////////////
'// イベント： フォーム アクティベート時
Private Sub UserForm_Activate()
'On Error GoTo ErrorHandler
    Dim shp     As Shape
    Dim spacer  As Integer
    
    Debug.Print ("activate")
    If VarType(Selection) = vbObject Then
    Debug.Print ("object selected:")
    Debug.Print (Selection.ShapeRange.Count)
    End If
    
    For Each shp In Selection.ShapeRange
        Select Case shp.AutoShapeType
            Case msoShapePie, msoShapeBlockArc
                 
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
    Debug.Print ("error")
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
            Case msoShapePie, msoShapeBlockArc, msoShapeCircularArrow
                shp.Rotation = 0
        End Select
    Next
End Sub



'// スピンボタン(開始角度)
Private Sub spnStart_SpinUp()
    angleStart = Int(angleStart / 15) * 15 + 15
    txtStart.Value = angleStart
    
    Call adjustArch
End Sub


Private Sub spnStart_SpinDown()
    If angleStart = Int(angleStart / 15) * 15 Then
        angleStart = Int(angleStart / 15) * 15 - 15
    Else
        angleStart = Int(angleStart / 15) * 15
    End If
    
    txtStart.Value = angleStart
    Call adjustArch
End Sub


'// スピンボタン(終了角度)
Private Sub spnEnd_SpinUp()
    angleEnd = Int(angleEnd / 15) * 15 + 15
    txtEnd.Value = angleEnd
    
    Call adjustArch
End Sub

Private Sub spnEnd_SpinDown()
    If angleEnd = Int(angleEnd / 15) * 15 Then
        angleEnd = Int(angleEnd / 15) * 15 - 15
    Else
        angleEnd = Int(angleEnd / 15) * 15
    End If
    
    txtEnd.Value = angleEnd
    Call adjustArch
End Sub






Private Sub adjustArch()
    Dim spacer  As Integer  '// シェイプごとの開始位置補正値を保持
    Dim shp     As Shape
    
    For Each shp In Selection.ShapeRange
        Select Case shp.AutoShapeType
            Case msoShapePie, msoShapeBlockArc
                spacer = -90
                shp.Adjustments.Item(1) = angleStart + spacer
                shp.Adjustments.Item(2) = angleEnd + spacer
            Case msoShapeCircularArrow
                spacer = -90
                shp.Adjustments.Item(4) = angleStart + spacer
                shp.Adjustments.Item(3) = angleEnd + spacer - shp.Adjustments.Item(2)
        End Select
    Next
    
End Sub



