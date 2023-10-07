Attribute VB_Name = "mdlGroup"
'// ////////////////////////////////////////////////////////////////////////////
'// プロジェクト   : 拡張ツール
'// タイトル       : グループ処理
'// モジュール     : mdlGroup
'// 説明           : 列・行のグループ処理、選択範囲の値の処理
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
Public Sub ribbonCallback_Group(control As IRibbonControl)
    Select Case control.ID
        Case "groupRow"                     '// グループ化 行
            Call psSetGroup_Row
        Case "groupCol"                     '// グループ化 列
            Call psSetGroup_Col
        Case "removeDup"                    '// 重複のカウント
            Call psDistinctVals
        Case "listDup"                      '// 重複を階層風に補正
            Call psGroupVals
    End Select
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// メソッド：   グループ設定（行）
'// 説明：       グループを自動設定する。
'// 引数：       なし
'// 戻り値：     なし
'// ////////////////////////////////////////////////////////////////////////////
Private Sub psSetGroup_Row()
    Dim idxStart    As Long
    Dim idxEnd      As Long
    Dim idxRow      As Long
    Dim idxCol      As Integer
    Dim tRange      As udTargetRange
    Dim childRange  As Range
    
    '// 事前チェック（アクティブシート保護、選択タイプ＝セル）
    If Not gfPreCheck(protectCont:=True, selType:=TYPE_RANGE) Then
        Exit Sub
    End If
    
    Call gsSuppressAppEvents
    
    '// アウトラインの集計位置を変更
    ActiveSheet.Outline.SummaryRow = xlAbove
    
    '// グループ設定
    For Each childRange In Selection.Areas
        tRange = gfGetTargetRange(ActiveSheet, childRange)
        
        idxStart = 0
        idxEnd = 0
        idxCol = tRange.minCol
        
        For idxRow = tRange.minRow To tRange.maxRow
            If idxStart = 0 Then
                idxStart = idxRow + 1
                idxEnd = idxRow + 1
            ElseIf Trim(Cells(idxRow, idxCol).Text) = BLANK Then
                idxEnd = idxRow
            ElseIf Trim(Cells(idxRow - 1, idxCol).Text) = BLANK Then
                Range(Cells(idxStart, 1), Cells(idxEnd, 1)).Rows.Group
                idxStart = idxRow + 1
                idxEnd = idxRow + 1
            Else
                idxStart = idxRow + 1
                idxEnd = idxRow + 1
            End If
        Next
        If idxStart < idxEnd Then
            Range(Cells(idxStart, 1), Cells(idxEnd, 1)).Rows.Group
        End If
      Next
      
      Call gsResumeAppEvents
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// メソッド：   グループ設定（列）
'// 説明：       グループを自動設定する。
'// 引数：       なし
'// 戻り値：     なし
'// ////////////////////////////////////////////////////////////////////////////
Private Sub psSetGroup_Col()
    Dim idxStart    As Long
    Dim idxEnd      As Long
    Dim idxRow      As Long
    Dim idxCol      As Integer
    Dim tRange      As udTargetRange
    Dim childRange  As Range
    
    '// 事前チェック（アクティブシート保護、選択タイプ＝セル）
    If Not gfPreCheck(protectCont:=True, selType:=TYPE_RANGE) Then
        Exit Sub
    End If
    
    Call gsSuppressAppEvents
    
    '// アウトラインの集計位置を変更
    ActiveSheet.Outline.SummaryColumn = xlLeft
    
    '// グループ設定
    For Each childRange In Selection.Areas
        tRange = gfGetTargetRange(ActiveSheet, childRange)
        
        idxStart = 0
        idxEnd = 0
        idxRow = tRange.minRow
        
        For idxCol = tRange.minCol To tRange.maxCol
            If idxStart = 0 Then
                idxStart = idxCol + 1
                idxEnd = idxCol + 1
            ElseIf Trim(Cells(idxRow, idxCol).Text) = BLANK Then
                idxEnd = idxCol
            ElseIf Trim(Cells(idxRow, idxCol - 1).Text) = BLANK Then
                Range(Cells(1, idxStart), Cells(1, idxEnd)).Columns.Group
                idxStart = idxCol + 1
                idxEnd = idxCol + 1
            Else
                idxStart = idxCol + 1
                idxEnd = idxCol + 1
            End If
        Next
        If idxStart < idxEnd Then
            Range(Cells(1, idxStart), Cells(1, idxEnd)).Columns.Group
        End If
    Next
    
    Call gsResumeAppEvents
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// メソッド：   値の重複を排除して一覧（カウント）
'// 説明：       行単位で重複値を排除する。
'//              複数列が選択された場合は、セルの値を不可読文字chr(127)で連結し、重複判定する
'// ////////////////////////////////////////////////////////////////////////////
Private Sub psDistinctVals()
    Dim idxRow      As Long
    Dim idxCol      As Integer
    Dim tRange      As udTargetRange
    
    Dim bff         As Variant
    Dim dict        As Object
    Dim keyString   As String
    Dim keyArray()  As String
    Dim resultSheet As Worksheet
    
    '// 事前チェック（選択タイプ＝セル）
    If Not gfPreCheck(selAreas:=1, selType:=TYPE_RANGE) Then
        Exit Sub
    End If
    
    Call gsSuppressAppEvents
    
    tRange = gfGetTargetRange(ActiveSheet, Selection)
    
    bff = Selection.Areas(1)
    Set dict = CreateObject("Scripting.Dictionary")
    
    For idxRow = 1 To tRange.maxRow - tRange.minRow + 1
        '// 行のセルを結合して文字列を一つに
        keyString = BLANK
        For idxCol = 1 To tRange.maxCol - tRange.minCol + 1
            If Not IsError(bff(idxRow, idxCol)) Then
                keyString = keyString & Chr(127) & bff(idxRow, idxCol)
            End If
        Next
        
        If Not dict.Exists(keyString) Then
            Call dict.Add(keyString, "1")
        Else
            dict.Item(keyString) = CStr(CLng(dict.Item(keyString)) + 1)
        End If
    Next
    
    '// 結果出力
    Call Workbooks.Add
    Set resultSheet = ActiveWorkbook.ActiveSheet
    
    '// ヘッダの設定。「カウント」のヘッダ位置を合わせるため、HDR_DISTINCT内の"@"を列数に合わせてReplaceする
    Call gsDrawResultHeader(resultSheet, Replace(HDR_DISTINCT, "@", String(tRange.Columns, ";")), 1)
    
    '// キーの配列をvariantに格納
    bff = dict.Keys
    
    For idxRow = 0 To dict.Count - 1
        keyArray = Split(bff(idxRow), Chr(127))  '// splitは添え字１から開始の仕様？
        For idxCol = 1 To UBound(keyArray)
            resultSheet.Cells(idxRow + 2, idxCol).Value = keyArray(idxCol)
        Next
        
        resultSheet.Cells(idxRow + 2, tRange.maxCol - tRange.minCol + 2).Value = dict.Item(bff(idxRow))
    Next
    
    '//フォント
    Call resultSheet.Cells.Select
    Selection.Font.Name = APP_FONT
    Selection.Font.Size = APP_FONT_SIZE
    
    '// 罫線描画
    Call gsPageSetup_Lines(resultSheet, 1)
    
    Set dict = Nothing
    Call gsResumeAppEvents
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// メソッド：   値を階層風に補正する
'// 説明：       重複値を階層風に補正する
'// ////////////////////////////////////////////////////////////////////////////
Private Sub psGroupVals()
    Dim idxRow        As Long
    Dim idxCol        As Integer
    Dim tRange        As udTargetRange
    Dim aryIdx        As Integer
    Dim aryLastVal(8) As String
    
    '// 事前チェック（選択エリア＝１、アクティブシート保護、選択タイプ＝セル）
    If Not gfPreCheck(selAreas:=1, protectCont:=True, selType:=TYPE_RANGE) Then
        Exit Sub
    End If
    
    '// チェック
    If Selection.Columns.Count > 8 Then
        Call MsgBox(MSG_TOO_MANY_COLS_8, vbOKOnly, APP_TITLE)
        Exit Sub
    End If
    
    Call gsSuppressAppEvents
    
    tRange = gfGetTargetRange(ActiveSheet, Selection)
    
    For idxRow = tRange.minRow To tRange.maxRow
        For idxCol = tRange.minCol To tRange.maxCol
            If (aryLastVal(idxCol - tRange.minCol) = BLANK) Or (aryLastVal(idxCol - tRange.minCol) <> Cells(idxRow, idxCol).Text) Then
                '// 直前の値が異なる場合 (含 直前の値がブランクの場合)
                '// 配下のレベルの直前の値をクリア
                For aryIdx = tRange.Columns To idxCol Step -1
                    aryLastVal(aryIdx - 1) = BLANK
                Next
                aryLastVal(idxCol - tRange.minCol) = Cells(idxRow, idxCol).Text
            Else
                Cells(idxRow, idxCol).Value = BLANK
            End If
        Next
    Next
    
    Call gsResumeAppEvents
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// END
'// ////////////////////////////////////////////////////////////////////////////

