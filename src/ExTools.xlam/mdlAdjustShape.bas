Attribute VB_Name = "mdlAdjustShape"
'// ////////////////////////////////////////////////////////////////////////////
'// プロジェクト   : 拡張ツール
'// タイトル       : オブジェクトの補正機能
'// モジュール     : mdlAdjustShape
'// 説明           : 鍵コネクタやブロック矢印などのオブジェクトの微調整機能
'//                  ※旧mdlFeatures（V2.1.1まで）
'// ////////////////////////////////////////////////////////////////////////////
'// Copyright (c) by Koichiro.
'// ////////////////////////////////////////////////////////////////////////////
Option Explicit
Option Base 0

'// ////////////////////////////////////////////////////////////////////////////
'// コンパイルスイッチ（"EXCEL" / "POWERPOINT"）
#Const OFFICE_APP = "EXCEL"


'// ////////////////////////////////////////////////////////////////////////////
'// メソッド：   リボンボタンコールバック管理(フォームなし)
'// 説明：       リボンからのコールバックをつかさどる
'//              押されたコントロールのIDを基に処理を呼び出す。
'// 引数：       control 対象コントロール
'// ////////////////////////////////////////////////////////////////////////////
Public Sub ribbonCallback_AdjustShape(control As IRibbonControl)
    Select Case control.ID
        Case "AdjShapeElbowConn"                                                '// 鍵コネクタの補正
            Call psAdjustElbowConnector
        Case "AdjShapeRoundRect"                                                '// 四角形の角丸み補正
            Call psAdjustRoundRect
        Case "AdjShapeBlockArrow"                                               '// ブロック矢印の傾き補正
            Call psAdjustBlockArrowHead
        Case "AdjShapeLine"                                                     '// 直線の傾き補正（0,45,90度）
            Call psAdjustLine
        Case "AdjShapeUngroup"                                                  '// 再帰でグループ解除
            Call psAdjustUngroup
        Case "AdjShapeOrderTile"                                                '// グリッドに整列
            Call psDistributeShapeGrid(0)
        Case "AdjShapeOrderTile_1"                                                '// グリッドに整列 1pt
            Call psDistributeShapeGrid(1)
        Case "AdjShapeOrderTile_2"                                                '// グリッドに整列 2pt
            Call psDistributeShapeGrid(2)
        Case "AdjShapeOrderTile_3"                                                '// グリッドに整列 3pt
            Call psDistributeShapeGrid(3)
        Case "AdjShapeOrderTile_4"                                                '// グリッドに整列 4pt
            Call psDistributeShapeGrid(4)
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
    
    '// 事前準備 //////////
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
    target = gfCeiling(target, 0.75)
    
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
Private Sub psAdjustBlockArrowHead()
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
                bff = gfMin2(.Height, .Width) * .Adjustments.Item(1)
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
                .Adjustments.Item(1) = target / gfMin2(.Height, .Width)
            End If
        End With
    Next
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// メソッド：   角の丸い四角形 丸み補正
'// 説明：       角の丸い四角形の丸みを、最もR（径）の小さいものに合わせる
'// ////////////////////////////////////////////////////////////////////////////
Private Sub psAdjustRoundRect()
    Dim target      As Double   '// 全ブロック矢印のAdjustment(1)をこのターゲットに合わせる。「短辺×Adjust値の最小値)」
    Dim bff         As Double
    Dim idx         As Integer
    
    '// 事前チェック（アクティブシート保護、選択タイプ＝シェイプ）
    If Not gfPreCheck(protectCont:=True, selType:=TYPE_SHAPE) Then
        Exit Sub
    End If
    
    '// ターゲット値(短辺×Adjust値の最小値)を取得 //////////
    target = 0
    For idx = 1 To ActiveWindow.Selection.ShapeRange.Count  '// shaperangeの開始インデックスは１から
        With ActiveWindow.Selection.ShapeRange(idx)
            If .AutoShapeType = msoShapeRoundedRectangle Then
                bff = gfMin2(.Height, .Width) * .Adjustments.Item(1)
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
                .Adjustments.Item(1) = target / gfMin2(.Height, .Width)
            End If
        End With
    Next
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// メソッド：   直線 角度補正
'// 説明：       直線の角度を、0,45,90度に補正する。元の位置の中心から回転させる
'// ////////////////////////////////////////////////////////////////////////////
Private Sub psAdjustLine()
    Dim idx         As Integer
    
    '// 事前チェック（アクティブシート保護、選択タイプ＝シェイプ）
    If Not gfPreCheck(protectCont:=True, selType:=TYPE_SHAPE) Then
        Exit Sub
    End If
    
    '// 角度設定
    For idx = 1 To ActiveWindow.Selection.ShapeRange.Count  '// shaperangeの開始インデックスは１から
        With ActiveWindow.Selection.ShapeRange(idx)
            If .Type = msoLine Then
                If .Width * .Height <> 0 Then
                    Select Case Atn(.Height / .Width) * 180 / (Atn(1) * 4)
                        Case Is <= 30   '// 0度に補正
                            .Top = IIf(.VerticalFlip, .Top - .Height / 2, .Top + .Height / 2)
                            .Height = 0
                        Case Is >= 70   '// 90度に補正
                            .Left = IIf(.VerticalFlip, .Left - .Width / 2, .Left + .Width / 2)
                            .Width = 0
                        Case Else   '// 45度に補正
                            If .Height > .Width Then
                                .Left = .Left - (.Height - .Width) / 2
                                .Width = .Height
                            Else
                                .Top = .Top - (.Width - .Height) / 2
                                .Height = .Width
                            End If
                    End Select
                End If
            End If
        End With
    Next
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// メソッド：   再帰でグループ解除
'// 説明：       ネストしたグループをすべて解除する。グループ解除部は _subに実装
'// ////////////////////////////////////////////////////////////////////////////
Private Sub psAdjustUngroup()
On Error GoTo ErrorHandler
    Dim idx         As Integer
'    Dim sh          As Shape
    
    '// 事前チェック（アクティブシート保護、選択タイプ＝シェイプ）
    If Not gfPreCheck(protectCont:=True, selType:=TYPE_SHAPE) Then
        Exit Sub
    End If
    
    For idx = 1 To ActiveWindow.Selection.ShapeRange.Count  '// shaperangeの開始インデックスは１から
        Call psAdjustUngroup_sub(ActiveWindow.Selection.ShapeRange(idx))
    Next
    Exit Sub

ErrorHandler:
#If OFFICE_APP = "EXCEL" Then
    Call gsResumeAppEvents
#End If
    Call gsShowErrorMsgDlg("psAdjustUngroup", Err)
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// メソッド：   再帰でグループ解除
'// 説明：       グループ解除実装部
'// ////////////////////////////////////////////////////////////////////////////
Private Sub psAdjustUngroup_sub(targetShape As Shape)
    Dim sh As Shape
    
    If targetShape.Type = msoGroup Then
        For Each sh In targetShape.Ungroup
            Call psAdjustUngroup_sub(sh)
        Next
    End If
End Sub



'// ////////////////////////////////////////////////////////////////////////////
'// メソッド：   グリッド整列
'// 説明：       メイン処理
'// ////////////////////////////////////////////////////////////////////////////
Private Sub psDistributeShapeGrid(spacing As Integer)
'On Error GoTo ErrorHandler
    Dim tls             As Shape    '// Top-Left-Shape. 左上の基準とするシェイプ
    Dim allShapes()     As Shape    '// すべてのシェイプを格納
    Dim rowHeader()     As Shape    '// 行ヘッダ（縦軸）のシェイプを格納
    Dim colHeader()     As Shape    '// 列ヘッダ（横軸）のシェイプを格納
    
    '// 事前チェック（アクティブシート保護、選択タイプ＝シェイプ）
    If Not gfPreCheck(protectCont:=True, selType:=TYPE_SHAPE) Then
        Exit Sub
    End If
    
#If OFFICE_APP = "EXCEL" Then
    allShapes = pfGetAllShapes(Selection.ShapeRange)                '// 全シェイプを配列に格納
    Set tls = pfGetTopLeftObject(Selection.ShapeRange)              '// TopLeftを取得
#ElseIf OFFICE_APP = "POWERPOINT" Then
    allShapes = pfGetAllShapes(ActiveWindow.Selection.ShapeRange)   '// 全シェイプを配列に格納
    Set tls = pfGetTopLeftObject(ActiveWindow.Selection.ShapeRange) '// TopLeftを取得
#End If
    
    '// 行ヘッダにあたるシェイプの配列を設定
    rowHeader = pfGetRowHeader(tls, allShapes, spacing)
    colHeader = pfGetColHeader(tls, allShapes, spacing)
    
    Call psAdjustAllShapes(allShapes, rowHeader, colHeader)
    Exit Sub
    
ErrorHandler:
#If OFFICE_APP = "EXCEL" Then
    Call gsResumeAppEvents
#End If
'    Call gsShowErrorMsgDlg("psDistributeShapeGrid", Err)
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// メソッド：   グリッド整列
'// 説明：       選択されたシェイプを全て配列に格納
'// ////////////////////////////////////////////////////////////////////////////
Public Function pfGetAllShapes(rng As ShapeRange) As Shape()
    Dim shp         As Shape
    Dim rslt()      As Shape
    Dim i           As Integer
    
    ReDim rslt(0)
    For Each shp In rng
        If rslt(0) Is Nothing Then
            Set rslt(0) = shp
        Else
            ReDim Preserve rslt(UBound(rslt) + 1)
            Set rslt(UBound(rslt)) = shp
        End If
    Next
    
    pfGetAllShapes = rslt
End Function


'// ////////////////////////////////////////////////////////////////////////////
'// メソッド：   グリッド整列
'// 説明：       基準となる、TopLeft位置のシェイプを取得
'// ////////////////////////////////////////////////////////////////////////////
Public Function pfGetTopLeftObject(rng As ShapeRange) As Shape
    Dim shp         As Shape
    Dim rslt        As Shape
    
    Set rslt = rng(1)
    
    '// Topが最も小さいシェイプを取得
    For Each shp In rng
        If shp.Top < rslt.Top Then
            Set rslt = shp
        End If
    Next
    
    '// 最小Topのシェイプの下辺よりもTopが小さく、かつ最小のLeftをもつシェイプを取得
    For Each shp In rng
        If shp.Top < (rslt.Top + rslt.Height) And shp.Left < rslt.Left Then
            Set rslt = shp
        End If
    Next
    
    Set pfGetTopLeftObject = rslt
    
'//　赤にする
'rslt.Fill.ForeColor.ObjectThemeColor = msoThemeColorAccent2
End Function


'// ////////////////////////////////////////////////////////////////////////////
'// メソッド：   グリッド整列
'// 説明：       行ヘッダ（縦軸）取得
'// ////////////////////////////////////////////////////////////////////////////
Private Function pfGetRowHeader(tls As Shape, ary() As Shape, spacing As Integer) As Shape()
'On Error GoTo ErrorHandler
'    Dim shp         As Shape
    Dim rslt()      As Shape
    Dim i           As Integer
    Dim bff         As Shape
    Dim idxS1       As Long
    Dim idxS2       As Long
    Dim heightTotal As Long     '// ヘッダのシェイプの高さの合計

    '// 縦軸に該当するオブジェクトを配列に格納
    ReDim rslt(0)
    For i = 0 To UBound(ary)
        If ary(i).Left < (tls.Left + (tls.Width * 0.9)) Then
            If Not rslt(0) Is Nothing Then
                ReDim Preserve rslt(UBound(rslt) + 1)
            End If
            Set rslt(UBound(rslt)) = ary(i)
        End If
    Next
    
    '// ソート
    idxS1 = 0
    ' 全テーブルの前からのループ
    Do While idxS1 < UBound(rslt)
        idxS2 = UBound(rslt)
        ' 終端から現在位置手前までのループ
        Do While idxS2 > idxS1
            ' 差し替え判定
            If rslt(idxS2).Top < rslt(idxS1).Top Then
                ' 差し替え
                Set bff = rslt(idxS2)
                Set rslt(idxS2) = rslt(idxS1)
                Set rslt(idxS1) = bff
            End If
            ' 前へ
            idxS2 = idxS2 - 1
        Loop
        ' 次へ
        idxS1 = idxS1 + 1
    Loop
    
    '// 選択解除
#If OFFICE_APP = "EXCEL" Then
    tls.TopLeftCell.Select
#ElseIf OFFICE_APP = "POWERPOINT" Then
    Call ActiveWindow.Selection.Unselect
#End If
    
    For i = 0 To UBound(rslt)
        Call rslt(i).Select(Replace:=False)
        heightTotal = heightTotal + rslt(i).Height
'        rslt(i).Line.ForeColor.RGB = vbRed
    Next
    
    '// オブジェクトが重なっている場合（高さの合計が最後のオブジェクトの終点よりも小さい）は、配置を広げる
    If UBound(rslt) > 0 Then
        If heightTotal >= (rslt(UBound(rslt)).Top + rslt(UBound(rslt)).Height) - tls.Top _
              Or spacing > 0 Then
'            rslt(UBound(rslt)).Line.ForeColor.RGB = vbRed
            rslt(UBound(rslt)).Top = tls.Top + heightTotal - rslt(UBound(rslt)).Height + UBound(rslt) * spacing
        End If
    End If
    
    
    If UBound(rslt) > 1 Then    '// 整列（Distribute）は３つ以上のオブジェクトが無いとエラーになるため
#If OFFICE_APP = "EXCEL" Then
        Call Selection.ShapeRange.Distribute(msoDistributeVertically, False)
#ElseIf OFFICE_APP = "POWERPOINT" Then
        Call ActiveWindow.Selection.ShapeRange.Distribute(msoDistributeVertically, False)
#End If
    End If
    
    pfGetRowHeader = rslt
    Exit Function
    
ErrorHandler:
'    Call gsShowErrorMsgDlg("pfGetRowHeader", Err)
End Function


'// ////////////////////////////////////////////////////////////////////////////
'// メソッド：   グリッド整列
'// 説明：       列ヘッダ（横軸）取得
'// ////////////////////////////////////////////////////////////////////////////
Private Function pfGetColHeader(tls As Shape, ary() As Shape, spacing As Integer) As Shape()
'On Error GoTo ErrorHandler
'    Dim shp         As Shape
    Dim rslt()      As Shape
    Dim i           As Integer
    Dim bff         As Shape
    Dim idxS1       As Long
    Dim idxS2       As Long
    Dim widthTotal  As Long     '// ヘッダのシェイプの幅の合計
    
    '// 横軸に該当するオブジェクトを配列に格納
    widthTotal = 0
    ReDim rslt(0)
    For i = 0 To UBound(ary)
        If ary(i).Top < (tls.Top + (tls.Height * 0.9)) Then
            If Not rslt(0) Is Nothing Then
                ReDim Preserve rslt(UBound(rslt) + 1)
            End If
            Set rslt(UBound(rslt)) = ary(i)
        End If
    Next
    
    '// ソート
    idxS1 = 0
    Do While idxS1 < UBound(rslt)                       '// 前からのループ
        idxS2 = UBound(rslt)
        Do While idxS2 > idxS1                          '// 終端から現在位置手前までのループ
            If rslt(idxS2).Left < rslt(idxS1).Left Then '// ソート入れ替え判定
                Set bff = rslt(idxS2)
                Set rslt(idxS2) = rslt(idxS1)
                Set rslt(idxS1) = bff
            End If
            idxS2 = idxS2 - 1
        Loop
        idxS1 = idxS1 + 1
    Loop
    
    '// 選択解除
#If OFFICE_APP = "EXCEL" Then
    tls.TopLeftCell.Select
#ElseIf OFFICE_APP = "POWERPOINT" Then
    Call ActiveWindow.Selection.Unselect
#End If

    For i = 0 To UBound(rslt)
        Call rslt(i).Select(Replace:=False)
        widthTotal = widthTotal + rslt(i).Width
'        rslt(i).Line.ForeColor.RGB = vbBlue
    Next
    
    '// オブジェクトが重なっている場合（幅の合計が最後のオブジェクトの終点よりも小さい）は、配置を広げる
    If UBound(rslt) > 0 Then
        If widthTotal >= (rslt(UBound(rslt)).Left + rslt(UBound(rslt)).Width) - tls.Left _
              Or spacing > 0 Then
'            rslt(UBound(rslt)).Line.ForeColor.RGB = vbBlue
            rslt(UBound(rslt)).Left = tls.Left + widthTotal - rslt(UBound(rslt)).Width + UBound(rslt) * spacing
        End If
    End If
    
    If UBound(rslt) > 1 Then    '// 整列（Distribute）は３つ以上のオブジェクトが無いとエラーになるため
#If OFFICE_APP = "EXCEL" Then
        Call Selection.ShapeRange.Distribute(msoDistributeHorizontally, False)
#ElseIf OFFICE_APP = "POWERPOINT" Then
        Call ActiveWindow.Selection.ShapeRange.Distribute(msoDistributeHorizontally, False)
#End If
    End If
    
    pfGetColHeader = rslt
    Exit Function
    
ErrorHandler:
    Call gsShowErrorMsgDlg("pfGetColHeader", Err)
End Function


'// ////////////////////////////////////////////////////////////////////////////
'// 全シェイプの配置
Private Sub psAdjustAllShapes(allShapes() As Shape, rowHeader() As Shape, colHeader() As Shape)
    Dim idx                 As Integer
    Dim idxHead             As Integer
    Dim bff                 As Double   '// 整列対象シェイプの中央位置を格納
    
    '// 全シェイプでのループ
    For idx = 0 To UBound(allShapes)
        '// 行ヘッダ（縦軸）でのループ
        For idxHead = 0 To UBound(rowHeader)
            bff = allShapes(idx).Top + allShapes(idx).Height / 2    '// 対象オブジェクトの中央ポジション（縦）
'            If bff >= allShapes(idx).Top And bff <= rowHeader(idxHead).Top + rowHeader(idxHead).Height Then
            If bff >= rowHeader(idxHead).Top And bff <= rowHeader(idxHead).Top + rowHeader(idxHead).Height Then
                allShapes(idx).Top = rowHeader(idxHead).Top
                allShapes(idx).Height = rowHeader(idxHead).Height
                Exit For
            End If
        Next
        
        '// 列ヘッダ（横軸）でのループ
        For idxHead = 0 To UBound(colHeader)
            bff = allShapes(idx).Left + allShapes(idx).Width / 2    '// 対象オブジェクトの中央ポジション（縦）
'            If bff >= allShapes(idx).Left And bff <= colHeader(idxHead).Left + colHeader(idxHead).Width Then
            If bff >= colHeader(idxHead).Left And bff <= colHeader(idxHead).Left + colHeader(idxHead).Width Then
                allShapes(idx).Left = colHeader(idxHead).Left
                allShapes(idx).Width = colHeader(idxHead).Width
                Exit For
            End If
        Next
        
        '// 選択解除をもとに戻す
        Call allShapes(idx).Select(Replace:=False)
    Next
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// END
'// ////////////////////////////////////////////////////////////////////////////


