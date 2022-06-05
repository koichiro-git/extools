Attribute VB_Name = "mdlFeatures"
'// ////////////////////////////////////////////////////////////////////////////
'// プロジェクト   : 拡張ツール
'// タイトル       : 共通関数
'// モジュール     : mdlFeatures
'// 説明           : フォームを持たない機能を記載
'//                  ※mdlCommonの肥大化に伴うファイル分割
'//                  ※個々のボタンから直接指定される仕組みにできる機能をこのモジュールに記載する
'// ////////////////////////////////////////////////////////////////////////////
'// Copyright (c) by Koichiro.
'// ////////////////////////////////////////////////////////////////////////////
Option Explicit
Option Base 0


'// ////////////////////////////////////////////////////////////////////////////
'// メソッド：   リボンボタンコールバック管理(フォームなし)
'// 説明：       リボンからのコールバックをつかさどる
'//              押されたコントロールのIDを基に処理を呼び出す。
'// 引数：       control 対象コントロール
'// ////////////////////////////////////////////////////////////////////////////
Public Sub ribbonCallback_f1(control As IRibbonControl)
    Select Case control.ID
        Case "AdjShapeElbowConn"                                                '// 鍵コネクタの補正
            Call psAdjustElbowConnector
        Case "AdjShapeRoundRect"                                                '// 四角形の角丸み補正
            Call psAdjustRoundRect
        Case "AdjShapeBlockArrow"                                               '// ブロック矢印の傾き補正
            Call psAdjustBlockArrowHead
    End Select
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// メソッド：   鍵コネクタ補正
'// 説明：       トーナメント表の鍵コネクタの補正位置を合わせる
'// ////////////////////////////////////////////////////////////////////////////
Private Sub psAdjustElbowConnector()
On Error GoTo ErrorHandler
    Dim topObjName  As String   '// トーナメントの頂上オブジェクト名
    Dim target      As Double   '// 全コネクタのAdjustment(1)をこのターゲットに合わせる。「コネクタ幅×Adjust値の最小値)」頂上オブジェクトに最も近い値を採用する。
    Dim idx         As Integer
    Dim elbows()    As Shape    '// 鍵コネクタのみを格納する配列
    Dim cntElbow    As Integer
    Dim bff         As Double
    
    '// 事前準備　//////////
    '// 事前チェック（アクティブシート保護、選択タイプ＝シェイプ）
    If Not gfPreCheck(protectCont:=True, selType:=TYPE_SHAPE) Then
        Exit Sub
    End If
    
    '// 鍵コネクタを取得
    cntElbow = 0
    For idx = 1 To ActiveWindow.Selection.ShapeRange.Count         '// shaperangeの開始インデックスは１から
        If ActiveWindow.Selection.ShapeRange(idx).Connector = msoTrue Then     '// ConnectorFormatは事前に参照可能か不明なため、If分岐をネスト
            If ActiveWindow.Selection.ShapeRange(idx).ConnectorFormat.Type = msoConnectorElbow Then
                ReDim Preserve elbows(cntElbow)
                Set elbows(cntElbow) = ActiveWindow.Selection.ShapeRange(idx)
                cntElbow = cntElbow + 1
            End If
        End If
    Next
    
    '// 最低２つ以上のコネクタが必要。ない場合はエラー
    If cntElbow < 2 Then
        Call MsgBox(MSG_SHAPE_MULTI_SELECT, vbOKCancel, APP_TITLE)
        Exit Sub
    End If
    
    '// 最初の2つのコネクタの連結オブジェクトを比較し、トーナメントの頂点オブジェクト名を取得
    If elbows(0).ConnectorFormat.BeginConnectedShape.Name = elbows(1).ConnectorFormat.BeginConnectedShape.Name Or _
        elbows(0).ConnectorFormat.BeginConnectedShape.Name = elbows(1).ConnectorFormat.EndConnectedShape.Name Then
        topObjName = elbows(0).ConnectorFormat.BeginConnectedShape.Name
    Else
        topObjName = elbows(0).ConnectorFormat.EndConnectedShape.Name
    End If
    
    '// ターゲット値(コネクタ幅×Adjust値の最小値)を取得　//////////
    target = 0
    For idx = 0 To UBound(elbows)
        With elbows(idx)
            If .ConnectorFormat.BeginConnectedShape.Name = topObjName Then
                bff = .Width * .Adjustments.Item(1)
            Else
                bff = .Width * (1 - .Adjustments.Item(1))
            End If
            
            If target = 0 Then
                target = bff
            ElseIf target > bff Then
                target = bff
            End If
        End With
    Next
    target = Application.WorksheetFunction.Ceiling(target, 0.75)
    
    '// 最小値に合わせてコネクタを設定
    For idx = 0 To UBound(elbows)
        With elbows(idx)
            If .ConnectorFormat.BeginConnectedShape.Name = topObjName Then
                .Adjustments.Item(1) = target / .Width
            Else
                .Adjustments.Item(1) = 1 - (target / .Width)
            End If
        End With
    Next
    
    Exit Sub
ErrorHandler:
    '//
    
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// メソッド：   ブロック矢印の先端角度補正
'// 説明：       ブロック矢印の先端角を、最も鈍角なものに合わせる
'// ////////////////////////////////////////////////////////////////////////////
Public Sub psAdjustBlockArrowHead()
    Dim target      As Double   '// 全ブロック矢印のAdjustment(1)をこのターゲットに合わせる。「短辺×Adjust値の最小値)」
    Dim bff         As Double
    Dim idx         As Integer
    
    '// 事前チェック（アクティブシート保護、選択タイプ＝シェイプ）
    If Not gfPreCheck(protectCont:=True, selType:=TYPE_SHAPE) Then
        Exit Sub
    End If
    
    '// ターゲット値(短辺×Adjust値の最小値)を取得　//////////
    target = 0
    For idx = 1 To ActiveWindow.Selection.ShapeRange.Count  '// shaperangeの開始インデックスは１から
        With ActiveWindow.Selection.ShapeRange(idx)
            If .AutoShapeType = msoShapePentagon Or _
                .AutoShapeType = msoShapeChevron Then
                bff = WorksheetFunction.Min(.Height, .Width) * .Adjustments.Item(1)
                If target = 0 Then
                    target = bff
                ElseIf target > bff Then
                    target = bff
                End If
            End If
        End With
    Next
    
    '// 最小値に合わせてブロック矢印の矢じりを設定
    For idx = 1 To ActiveWindow.Selection.ShapeRange.Count  '// shaperangeの開始インデックスは１から
        With ActiveWindow.Selection.ShapeRange(idx)
            If .AutoShapeType = msoShapePentagon Or _
                .AutoShapeType = msoShapeChevron Then
                .Adjustments.Item(1) = target / WorksheetFunction.Min(.Height, .Width)
            End If
        End With
    Next
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// メソッド：   角の丸い四角形 丸み補正
'// 説明：       角の丸い四角形の丸みを、最もR（径）の小さいものに合わせる
'// ////////////////////////////////////////////////////////////////////////////
Public Sub psAdjustRoundRect()
    Dim target      As Double   '// 全ブロック矢印のAdjustment(1)をこのターゲットに合わせる。「短辺×Adjust値の最小値)」
    Dim bff         As Double
    Dim idx         As Integer
    
    '// 事前チェック（アクティブシート保護、選択タイプ＝シェイプ）
    If Not gfPreCheck(protectCont:=True, selType:=TYPE_SHAPE) Then
        Exit Sub
    End If

    '// ターゲット値(短辺×Adjust値の最小値)を取得　//////////
    target = 0
    For idx = 1 To ActiveWindow.Selection.ShapeRange.Count  '// shaperangeの開始インデックスは１から
        With ActiveWindow.Selection.ShapeRange(idx)
            If .AutoShapeType = msoShapeRoundedRectangle Then
                bff = WorksheetFunction.Min(.Height, .Width) * .Adjustments.Item(1)
                If target = 0 Then
                    target = bff
                ElseIf target > bff Then
                    target = bff
                End If
            End If
        End With
    Next
    
    '// 最小値に合わせて四角形の角を設定
    For idx = 1 To ActiveWindow.Selection.ShapeRange.Count  '// shaperangeの開始インデックスは１から
        With ActiveWindow.Selection.ShapeRange(idx)
            If .AutoShapeType = msoShapeRoundedRectangle Then
                .Adjustments.Item(1) = target / WorksheetFunction.Min(.Height, .Width)
            End If
        End With
    Next
    
End Sub
    


