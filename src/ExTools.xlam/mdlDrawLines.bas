Attribute VB_Name = "mdlDrawLines"
'// ////////////////////////////////////////////////////////////////////////////
'// プロジェクト   : 拡張ツール
'// タイトル       : 罫線描画機能
'// モジュール     : mdlDrawLines
'// 説明           : 行・列のヘッダおよびデータ部分の罫線を描画
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
Public Sub ribbonCallback_DrawLines(control As IRibbonControl)
    Select Case control.ID
        Case "BorderRowHead"                '// 行ヘッダの罫線
            Call gsDrawLine_Header
        Case "BorderColHead"                '// 列ヘッダの罫線
            Call gsDrawLine_Header_Vert
        Case "BorderData"                   '// データ領域の罫線
            Call gsDrawLine_Data
    End Select
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// メソッド：   罫線描画（ヘッダ）
'// 説明：       ヘッダ部の罫線を描画する（横）
'// ////////////////////////////////////////////////////////////////////////////
Public Sub gsDrawLine_Header()
    Dim baseRow As Long     '// 選択領域の開始位置
    Dim baseCol As Integer  '// 選択領域の開始位置
    Dim selRows As Long     '// 選択領域の行数
    Dim selCols As Integer  '// 選択領域の列数
    Dim idxRow  As Long
    Dim idxCol  As Integer
    Dim offRow  As Long
    Dim offCol  As Integer
    Dim childRange As Range
    
    '// 事前チェック（アクティブシート保護、選択タイプ＝セル）
    If Not gfPreCheck(protectCont:=True, selType:=TYPE_RANGE) Then
        Exit Sub
    End If
    
    Call gsSuppressAppEvents
    
    For Each childRange In Selection.Areas
        '// 罫線をクリア
        childRange.Borders.LineStyle = xlNone
        childRange.Borders(xlDiagonalDown).LineStyle = xlNone
        childRange.Borders(xlDiagonalUp).LineStyle = xlNone
        
        '// 選択範囲の開始・終了位置取得
        baseRow = childRange.Row
        baseCol = childRange.Column
        selRows = childRange.Rows.Count
        selCols = childRange.Columns.Count
        
        For idxRow = baseRow To baseRow + selRows
            For idxCol = baseCol To baseCol + selCols
                offRow = 0
                offCol = 0
                If (Cells(idxRow, idxCol).Text <> BLANK) Or ((idxRow = baseRow) And (idxCol = baseCol)) Then
                    For offRow = idxRow To baseRow + selRows - 1
                        If (offRow = idxRow) Or Cells(offRow, idxCol).Value = BLANK Then
                            Cells(offRow, idxCol).Borders(xlEdgeLeft).LineStyle = xlContinuous
                        Else
                            Exit For
                        End If
                    Next
                    For offCol = idxCol To baseCol + selCols - 1
                        If (offCol = idxCol) Or Cells(idxRow, offCol).Text = BLANK Then
                            Cells(idxRow, offCol).Borders(xlEdgeTop).LineStyle = xlContinuous
                            If Cells(idxRow, offCol).Borders(xlEdgeRight).LineStyle = xlContinuous Then
                                Exit For
                            End If
                        Else
                            Exit For
                        End If
                    Next
                End If
                '// 最大列に達した場合は終了
                If idxCol = Columns.Count Then
                    Exit For
                End If
            Next
            '// 最大行に達した場合は終了
            If idxRow = Rows.Count Then
                Exit For
            End If
        Next
        
        With childRange.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With
        With childRange.Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With
    Next
    
    Call gsResumeAppEvents
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// メソッド：   罫線描画（ヘッダ）：縦
'// 説明：       ヘッダ部の罫線を描画する（縦）
'// 引数：       なし
'// 戻り値：     なし
'// ////////////////////////////////////////////////////////////////////////////
Public Sub gsDrawLine_Header_Vert()
    Dim baseRow As Long     '// 選択領域の開始位置
    Dim baseCol As Integer  '// 選択領域の開始位置
    Dim selRows As Long     '// 選択領域の行数
    Dim selCols As Integer  '// 選択領域の列数
    Dim idxRow  As Long
    Dim idxCol  As Integer
    Dim offRow  As Long
    Dim offCol  As Integer
    Dim childRange As Range
  
    '// 事前チェック（アクティブシート保護、選択タイプ＝セル）
    If Not gfPreCheck(protectCont:=True, selType:=TYPE_RANGE) Then
        Exit Sub
    End If
    
    Call gsSuppressAppEvents
    
    For Each childRange In Selection.Areas
        '// 罫線をクリア
        childRange.Borders.LineStyle = xlNone
        childRange.Borders(xlDiagonalDown).LineStyle = xlNone
        childRange.Borders(xlDiagonalUp).LineStyle = xlNone
        
        '// 選択範囲の開始・終了位置取得
        baseRow = childRange.Row
        baseCol = childRange.Column
        selRows = childRange.Rows.Count
        selCols = childRange.Columns.Count
      
        For idxCol = baseCol To baseCol + selCols
            For idxRow = baseRow To baseRow + selRows
                offRow = 0
                offCol = 0
                If (Cells(idxRow, idxCol).Value <> BLANK) Or ((idxRow = baseRow) And (idxCol = baseCol)) Then
                    For offCol = idxCol To baseCol + selCols - 1
                        If (offCol = idxCol) Or Cells(idxRow, offCol).Value = BLANK Then
                            Cells(idxRow, offCol).Borders(xlEdgeTop).LineStyle = xlContinuous
                        Else
                            Exit For
                        End If
                    Next
                    For offRow = idxRow To baseRow + selRows - 1
                        If (offRow = idxRow) Or Cells(offRow, idxCol).Value = BLANK Then
                            Cells(offRow, idxCol).Borders(xlEdgeLeft).LineStyle = xlContinuous
                            If Cells(offRow, idxCol).Borders(xlEdgeBottom).LineStyle = xlContinuous Then
                                Exit For
                            End If
                        Else
                            Exit For
                        End If
                    Next
                End If
                '// 最大行に達した場合は終了
                If idxRow = Rows.Count Then
                    Exit For
                End If
            Next
            '// 最大列に達した場合は終了
            If idxCol = Columns.Count Then
                Exit For
            End If
        Next
    
        With childRange.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With
        With childRange.Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With
    Next
    
    Call gsResumeAppEvents
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// メソッド：   罫線描画（データ）
'// 説明：       データ部の罫線を描画する
'//              選択範囲周辺部をxlThin、内部をxlHairlineで描画する
'// 引数：       なし
'// 戻り値：     なし
'// ////////////////////////////////////////////////////////////////////////////
Public Sub gsDrawLine_Data()
    '// 事前チェック（アクティブシート保護、選択タイプ＝セル）
    If Not gfPreCheck(protectCont:=True, selType:=TYPE_RANGE) Then
        Exit Sub
    End If
    
    Call gsSuppressAppEvents
    
    '// 罫線描画
    Selection.Borders.LineStyle = xlContinuous
    Selection.Borders.Weight = xlThin
    
    If Selection.Columns.Count > 1 Then
        With Selection.Borders(xlInsideVertical)
            .LineStyle = xlContinuous
            .Weight = xlHairline
        End With
    End If
    
    If Selection.Rows.Count > 1 Then
        With Selection.Borders(xlInsideHorizontal)
            .LineStyle = xlContinuous
            .Weight = xlHairline
        End With
    End If
    
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone

    Call gsResumeAppEvents
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// END
'// ////////////////////////////////////////////////////////////////////////////

