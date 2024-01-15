VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmFileList 
   Caption         =   "ファイル一覧出力"
   ClientHeight    =   4455
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6150
   OleObjectBlob   =   "frmFileList.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmFileList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'// ////////////////////////////////////////////////////////////////////////////
'// プロジェクト   : 拡張ツール
'// タイトル       : ファイル一覧出力フォーム
'// モジュール     : frmFileList
'// 説明           : シートの比較を行う
'// ////////////////////////////////////////////////////////////////////////////
'// Copyright (c) by Koichiro.
'// ////////////////////////////////////////////////////////////////////////////
Option Explicit
Option Base 0

'// ////////////////////////////////////////////////////////////////////////////
'// プライベート定数
Private Const pUNLIMITED_DEPTH    As Integer = "32767"

'// プライベート変数
Private pRootDir        As String   '// ディレクトリ取得開始位置
Private pExtentions()   As String   '// 拡張子の配列
Private pMaxDepth       As Integer  '// 最大深度


'// //////////////////////////////////////////////////////////////////
'// イベント： フォーム 初期化時
Private Sub UserForm_Initialize()
    '// コンボボックス設定
    Call gsSetCombo(cmbDirDepth, CMB_LST_DEPTH, 9)
    Call gsSetCombo(cmbTargetFile, CMB_LST_TARGET, 0)
    Call gsSetCombo(cmbFileSize, CMB_LST_SIZE, 0)
    
    '// キャプション設定
    frmFileList.Caption = LBL_LST_FORM
    cmdExecute.Caption = LBL_COM_EXEC
    cmdClose.Caption = LBL_COM_CLOSE
    cmdRootDir.Caption = LBL_COM_BROWSE
    ckbPath.Caption = LBL_LST_REL_PATH
    ckbHyperLink.Caption = LBL_COM_HYPERLINK
    lblRoot.Caption = LBL_LST_ROOT
    lblDepth.Caption = LBL_LST_DEPTH
    lblTarget.Caption = LBL_LST_TARGET
    lblExtentions.Caption = LBL_LST_EXT
    lblSize.Caption = LBL_LST_SIZE
End Sub


'// //////////////////////////////////////////////////////////////////
'// イベント： 閉じるボタン クリック時
Private Sub cmdClose_Click()
    Call Me.Hide
End Sub


'// //////////////////////////////////////////////////////////////////
'// イベント： 参照ボタン クリック時
Private Sub cmdRootDir_Click()
    Dim FilePath  As String
    
    If Not gfShowSelectFolder(0, FilePath) Then
        Exit Sub
    Else
        txtRootDir.Text = FilePath
    End If
End Sub


'// //////////////////////////////////////////////////////////////////
'// イベント： 対象ファイルコンボ 更新時
Private Sub cmbTargetFile_Change()
    Select Case cmbTargetFile.Value
        Case "0"
            txtExtentions.Enabled = False
            txtExtentions.BackColor = CLR_DISABLED
        Case Else
            txtExtentions.Enabled = True
            txtExtentions.BackColor = CLR_ENABLED
    End Select
End Sub


'// //////////////////////////////////////////////////////////////////
'// イベント： 実行ボタン クリック時
Private Sub cmdExecute_Click()
    If Trim(txtRootDir.Text) = BLANK Then  '//空白チェック
        Call MsgBox(MSG_NO_DIR, vbOKOnly, APP_TITLE)
        Call txtRootDir.SetFocus
    Else
        Call gsSuppressAppEvents
        Call psShowFileList
        Call gsResumeAppEvents
        Call Me.Hide
    End If
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// メソッド：   リスト出力メイン
'// 説明：       リスト出力を行う。
'// 引数：       なし
'// 戻り値：     なし
'// ////////////////////////////////////////////////////////////////////////////
Private Sub psShowFileList()
On Error GoTo ErrorHandler
    Dim fs          As Object
    Dim rootDir     As Object
    Dim wkSheet     As Worksheet
    Dim sizeUnit    As Double
    Dim sizeUnitTxt As String
    Dim sizeFormat  As String
    Dim idx         As Integer
  
    '// 設定値の記憶
    pRootDir = txtRootDir.Text                      '// ルートの設定
    pExtentions = Split(txtExtentions.Text, ";")    '// 拡張子 (trim処理要)
    For idx = 0 To UBound(pExtentions)
        pExtentions(idx) = LCase(Trim(pExtentions(idx)))
    Next
    pMaxDepth = CInt(cmbDirDepth.Value)             '// 最大深度
    If cmbFileSize.Value = 0 Then                   '// ファイルサイズ単位
        sizeUnit = 1
        sizeUnitTxt = "B"
        sizeFormat = "#,##0 "
    ElseIf cmbFileSize.Value = 1 Then
        sizeUnit = 1024
        sizeUnitTxt = "KB"
        sizeFormat = "#,##0.0_ "
    ElseIf cmbFileSize.Value = 2 Then
        sizeUnit = 1048576
        sizeUnitTxt = "MB"
        sizeFormat = "#,##0.0_ "
    End If
    
    '// ファイルシステムオブジェクトの作成と検索パス確認
    Set fs = CreateObject("Scripting.FileSystemObject")
    If Not fs.FolderExists(pRootDir) Then
        Call MsgBox(MSG_DIR_NOT_EXIST, vbOKOnly, APP_TITLE)
        Set fs = Nothing
        Exit Sub
    End If
 
    Call Workbooks.Add
    Set wkSheet = ActiveWorkbook.ActiveSheet
  
    '// ヘッダの描画
    Call gsDrawResultHeader(wkSheet, Replace(HDR_LST, "$", sizeUnitTxt), 1)
  
    '// ルートの出力
    Set rootDir = fs.GetFolder(pRootDir)
    wkSheet.Cells(2, 1).Value = rootDir.Path
    If Not rootDir.IsRootFolder Then
        wkSheet.Cells(2, 3).Value = rootDir.DateCreated
        wkSheet.Cells(2, 4).Value = rootDir.DateLastModified
    End If
    '// 塗りつぶし
    Range(wkSheet.Cells(2, 1), wkSheet.Cells(2, 8)).Interior.ColorIndex = COLOR_ROW
    
    '// ファイル出力ルーチンの呼び出し（再帰）
    Call psGetFileList(wkSheet, fs, rootDir.Path, 2, 0, 0, cmbTargetFile.Value, ckbHyperLink.Value, sizeUnit)
    Set fs = Nothing
  
    '// //////////////////////////////////////////////////////
    '// 書式の設定
    '// 列の書式
    wkSheet.Columns("A:B").NumberFormatLocal = "@"              '// パス、ファイル名
    wkSheet.Columns("C:D").NumberFormatLocal = "yyyy/mm/dd"     '// 作成日、更新日
    
    wkSheet.Columns("E").NumberFormatLocal = sizeFormat         '// ファイルサイズ
    
    '// 幅の設定
    wkSheet.Columns("A").ColumnWidth = 15
    wkSheet.Columns("B").ColumnWidth = 20
    wkSheet.Columns("C:D").ColumnWidth = 9
    
    '// 枠線の設定
    Call gsPageSetup_Lines(wkSheet, 1)
    
    '//フォント
    wkSheet.Cells.Font.Name = APP_FONT
    wkSheet.Cells.Font.Size = APP_FONT_SIZE
    
    '// ファイル属性説明記載
    wkSheet.Cells(1, 7).AddComment ("rhsa: Read only, Hidden, System file, Archive")
    
    '// 後処理
    Call wkSheet.Cells(1, 1).Select
    ActiveWorkbook.Saved = True
    Exit Sub

ErrorHandler:
    Call gsShowErrorMsgDlg("frmFileList.pfShowFileList", Err, Nothing)
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// メソッド：   ファイル一覧出力
'// 説明：       引数のディレクトリ以下の要素を出力する
'// 引数：       wkSheet: 出力対象シート
'//              fs: 対象ファイルシステムオブジェクト
'//              dirName: 対象ディレクトリ名
'//              idxRow: 結果出力行
'//              depth: ディレクトリ深度
'//              mode_Dir: ディレクトリ検索モード 0:全て,1:空を除外,2:空のみ
'//              mode_File: ファイル検索モード
'//              addLink: ハイパーリンク付加
'// 戻り値：     なし
'// ////////////////////////////////////////////////////////////////////////////
Private Sub psGetFileList(wkSheet As Worksheet, fs As Object, dirName As String, ByRef idxRow As Long, depth As Integer, mode_Dir As Integer, mode_File As Integer, addLink As Boolean, sizeUnit As Double)
On Error GoTo ErrorHandler
    Dim currentRow  As Long
    Dim parentDir   As Object
    Dim children    As Object
    Dim cnt         As Integer
    Dim isTarget    As Boolean
    Dim isEmptyDir  As Boolean
  
    currentRow = idxRow
  
    Set parentDir = fs.GetFolder(dirName)
    isEmptyDir = True
  
    '// ステータスバー更新
    Application.StatusBar = MSG_PROCESSING & " [ " & dirName & " ]"
    
    '// ファイルの出力
    For Each children In parentDir.files
        isEmptyDir = False
        With children
            Select Case mode_File
                Case "0"
                    isTarget = True
                Case "1"
                    isTarget = False
                    For cnt = 0 To UBound(pExtentions)
                        If LCase(Right(.Name, Len(pExtentions(cnt)))) = pExtentions(cnt) Then
                            isTarget = True
                            Exit For
                        End If
                    Next
                Case "2"
                    isTarget = True
                    For cnt = 0 To UBound(pExtentions)
                        If LCase(Right(.Name, Len(pExtentions(cnt)))) = pExtentions(cnt) Then
                            isTarget = False
                            Exit For
                        End If
                    Next
            End Select
            
            If isTarget Then
                idxRow = idxRow + 1
                wkSheet.Cells(idxRow, 2).Value = .Name
                wkSheet.Cells(idxRow, 3).Value = .DateCreated
                wkSheet.Cells(idxRow, 4).Value = .DateLastModified
                wkSheet.Cells(idxRow, 5).Value = .Size / sizeUnit
                wkSheet.Cells(idxRow, 6).Value = .Type
                wkSheet.Cells(idxRow, 7).Value = pfGetAttrString(.Attributes)
                
                '// ゼロバイトファイルの備考欄
                If .Size = 0 Then
                    wkSheet.Cells(idxRow, 8).Value = MSG_ZERO_BYTE
                End If
                '// リンクの設定
                If addLink Then
                    Call wkSheet.Cells(idxRow, 2).Hyperlinks.Add(Anchor:=wkSheet.Cells(idxRow, 2), Address:=.parentfolder & "\" & .Name)
                End If
            End If
        End With
    Next
  
    '// サブフォルダの出力
    For Each children In parentDir.SubFolders
        isEmptyDir = False
        idxRow = idxRow + 1
        With children
            wkSheet.Cells(idxRow, 1).Value = IIf(ckbPath.Value, "." & Mid(.Path, Len(pRootDir) + 1), .Path)
            wkSheet.Cells(idxRow, 3).Value = .DateCreated
            wkSheet.Cells(idxRow, 4).Value = .DateLastModified
            '// 塗りつぶし
            Range(wkSheet.Cells(idxRow, 1), wkSheet.Cells(idxRow, 8)).Interior.ColorIndex = COLOR_ROW
            
            '// リンクの設定
            If addLink Then
                Call wkSheet.Cells(idxRow, 1).Hyperlinks.Add(Anchor:=Cells(idxRow, 1), Address:=.Path)
            End If
            
            '// 子ディレクトリの再帰呼び出し
            If depth < pMaxDepth Then
                Call psGetFileList(wkSheet, fs, .Path, idxRow, depth + 1, mode_Dir, mode_File, addLink, sizeUnit)
            Else
                wkSheet.Cells(idxRow, 5).Value = fs.GetFolder(.Path).Size / sizeUnit    '// 配下のファイルサイズを取得
                wkSheet.Cells(idxRow, 8).Value = MSG_MAX_DEPTH
            End If
        End With
    Next

    '// 空フォルダの備考欄
    If isEmptyDir Then
        wkSheet.Cells(currentRow, 8).Value = MSG_EMPTY_DIR
    End If
    Exit Sub
  
ErrorHandler:
    If Err.Number = 70 Then
        wkSheet.Cells(currentRow, 8).Value = MSG_ERR_PRIV
    Else
        Call gsShowErrorMsgDlg("frmFileList.psGetFileList", Err, Nothing)
    End If
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// メソッド：   属性表示文字列生成
'// 説明：       引数の属性数値を文字列に変換する
'// 引数：       targetVal: 属性を表す整数値
'// 戻り値：     属性を表す文字列 "rhsa" 書式
'// ////////////////////////////////////////////////////////////////////////////
Private Function pfGetAttrString(targetVal As Integer)
    pfGetAttrString = IIf(targetVal And vbReadOnly, "r", "-") & IIf(targetVal And vbHidden, "h", "-") & IIf(targetVal And vbSystem, "s", "-") & IIf(targetVal And vbArchive, "a", "-")
End Function


'// ////////////////////////////////////////////////////////////////////////////
'// END
'// ////////////////////////////////////////////////////////////////////////////
