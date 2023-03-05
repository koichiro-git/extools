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
Option Explicit
Option Base 0



'// //////////////////////////////////////////////////////////////////
'// イベント： フォーム 初期化時
Private Sub UserForm_Initialize()
    cmdClose.Caption = LBL_COM_CLOSE
End Sub


'// //////////////////////////////////////////////////////////////////
'// イベント： フォーム アクティベート時
Private Sub UserForm_Activate()
'On Error GoTo ErrorHandler
    Dim spacer  As Integer  '// シェイプごとの開始位置補正値を保持
    
    Debug.Print ("activate")
    If VarType(Selection) = vbObject Then
    Debug.Print ("object selected:")
    Debug.Print (Selection.ShapeRange.Count)
    End If
    
'    For Each s In Selection.ShapeRange
'
'
'    Next
    
    Select Case Selection.ShapeRange(1).AutoShapeType
        Case msoShapePie, msoShapeBlockArc
            spacer = -90
            txtStart.Value = Int(Selection.ShapeRange.Adjustments.Item(1)) + spacer
            txtEnd.Value = Int(Selection.ShapeRange.Adjustments.Item(2)) + spacer
        Case msoShapeCircularArrow
            spacer = -90
            txtStart.Value = Int(Selection.ShapeRange.Adjustments.Item(4)) + spacer
            txtEnd.Value = Int(Selection.ShapeRange.Adjustments.Item(3)) + spacer
        Case Else
            Debug.Print Selection.ShapeRange(1).AutoShapeType
    End Select
    
    
    
    Exit Sub

ErrorHandler:
    Debug.Print ("error")
End Sub

'// //////////////////////////////////////////////////////////////////
'// イベント： 閉じるボタン クリック時
Private Sub cmdClose_Click()
    Call Me.Hide
End Sub


Private Sub cmdResetRotation_Click()
    Dim shp As Shape
    
    For Each shp In Selection.ShapeRange
        shp.Rotation = 0
    Next
End Sub

Private Sub spnEnd_SpinDown()
txtEnd.Value = txtEnd.Value - 1
End Sub

Private Sub spnEnd_SpinUp()
txtEnd.Value = txtEnd.Value + 1

End Sub

Private Sub spnStart_SpinDown()
    txtStart.Value = txtStart.Value - 1
End Sub

Private Sub spnStart_SpinUp()
    txtStart.Value = txtStart.Value + 1
End Sub



Private Sub txtStart_Change()
    Call adjustArch(Selection.ShapeRange(1), txtStart.Value, txtEnd.Value)
    spnStart.Value = 0
End Sub

Private Sub txtEnd_Change()
    Call adjustArch(Selection.ShapeRange(1), txtStart.Value, txtEnd.Value)
End Sub



Private Sub adjustArch(targetShape As Shape, startAngle As Integer, endAngle As Integer)
    Dim spacer  As Integer  '// シェイプごとの開始位置補正値を保持
    
    Select Case targetShape.AutoShapeType
        Case msoShapePie, msoShapeBlockArc
            spacer = -90
            targetShape.Adjustments.Item(1) = startAngle + spacer
            targetShape.Adjustments.Item(2) = endAngle + spacer
        Case msoShapeCircularArrow
            spacer = -90
            targetShape.Adjustments.Item(4) = startAngle + spacer
            targetShape.Adjustments.Item(3) = endAngle + spacer
    End Select
End Sub



