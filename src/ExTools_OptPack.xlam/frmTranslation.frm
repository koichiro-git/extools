VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTranslation 
   Caption         =   "翻訳"
   ClientHeight    =   1860
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4605
   OleObjectBlob   =   "frmTranslation.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmTranslation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'// ////////////////////////////////////////////////////////////////////////////
'// プロジェクト   : 拡張ツール 追加パック
'// タイトル       : 翻訳フォーム
'// モジュール     : frmTranslation
'// 説明           : 選択範囲のセル文字列を翻訳する
'// ////////////////////////////////////////////////////////////////////////////
'// Copyright (c) by Koichiro.
'// ////////////////////////////////////////////////////////////////////////////
Option Explicit
Option Base 0

'// ////////////////////////////////////////////////////////////////////////////
'// プライベート定数
Private Const URL_DEEPL_FREE = "https://api-free.deepl.com/v2/translate?"
Private Const URL_DEEPL_LICENSED = "https://api.deepl.com/v2/translate?"

Private Const SEPARATOR = "--" + vbLf
Private Const TAG_LINE_FEED = "<LF/>"

'// プライベート変数
Private auth_key        As String
Private license         As String
Private url_deepl       As String

'// //////////////////////////////////////////////////////////////////
'// イベント： フォーム 初期化時
Private Sub UserForm_Initialize()
    '// 設定ファイルからキーを読み込み
    auth_key = gfGetIniFileSetting("TRANSLATE", "DEEPL_AUTH_KEY")   '// 認証キー
    license = gfGetIniFileSetting("TRANSLATE", "DEEPL_LICENSE")     '// ライセンス（FREE/PRO）
    If license <> "FREE" Then
        url_deepl = URL_DEEPL_LICENSED
    Else
        url_deepl = URL_DEEPL_FREE
    End If
    
    '// コンボボックス設定
    Call gsSetCombo(cmbLanguage, CMB_TRN_LANGUAGE, 0)
    Call gsSetCombo(cmbOutput, CMB_TRN_OUTPUT, 0)

    '// キャプション設定
    frmTranslation.Caption = IIf(license <> "FREE", BLANK, LBL_TRN_FORM & " " & pfGetServiceUsage)
    lblLanguage.Caption = LBL_TRN_LANG
    lblOutput.Caption = LBL_TRN_OUTPUT
End Sub


'// //////////////////////////////////////////////////////////////////
'// イベント： 閉じるボタン クリック時
Private Sub cmdClose_Click()
    Call Me.Hide
End Sub


'// //////////////////////////////////////////////////////////////////
'// イベント： 実行ボタン クリック時
Private Sub cmdExecute_Click()
    Dim tCells      As Range    '// 変換対象セル
    
    '// 事前チェック（アクティブシート保護、選択タイプ＝セル）
    If Not gfPreCheck(protectCont:=True, selType:=TYPE_RANGE) Then
        Exit Sub
    End If
    Call gsSuppressAppEvents
    
    If Selection.Count > 1 Then
        Set tCells = Selection.SpecialCells(xlCellTypeConstants, xlNumbers + xlTextValues)
    Else
        Set tCells = ActiveCell
    End If
    
    Call psTranslate_DeepL(tCells)  '// DeepL
    
    Call gsResumeAppEvents
    Call Me.Hide
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// メソッド：   翻訳 主関数 (DeepL)
'// 説明：       選択範囲の文字列を翻訳する
'// 引数：       tCells 対象範囲
'// ////////////////////////////////////////////////////////////////////////////
Private Sub psTranslate_DeepL(tCells As Range)
'On Error GoTo ErrorHandler
    Dim httpReq         As New XMLHTTP60
    Dim htmlDoc         As New HTMLDocument
    Dim reqParam        As String
    Dim sourceText      As String
    Dim resultText      As String
    Dim appendIdx       As Integer
    
    Dim tCell           As Range
    
    For Each tCell In tCells
        sourceText = EncodeURL2(tCell.Text, htmlDoc)
        reqParam = "&target_lang=" & cmbLanguage.Value & _
                   "&auth_key=" & auth_key & _
                   "&text=" & sourceText & _
                   "&tag_handling=xml"
        
        Call httpReq.Open("POST", url_deepl)
        Call httpReq.setRequestHeader("Content-Type", "application/x-www-form-urlencoded")
        Call httpReq.send(reqParam)
        
        Application.StatusBar = tCell.Address & ": " & Left(tCell.Text, 30)
        Do While httpReq.readyState < 4
            DoEvents
        Loop
        
        '// レスポンス取得後（JSONライブラリは使用せず、文字列として処理する）
        If httpReq.Status = 200 Then
            '// JSONより結果文字列抽出。形式は {"translations":[{"detected_source_language":"EN","text":"★★★"}]}
            resultText = Mid(httpReq.responseText, InStr(1, httpReq.responseText, """text"":") + 8)     '// "text":" より後を取得
            resultText = Left(resultText, Len(resultText) - 4)  '// 最後の "}]} を削除
            resultText = DecodeResultText(resultText)
        Else
            resultText = "Error: HTTP Response=" & httpReq.Status
        End If
        
        '// 結果文字列を設定
        Select Case cmbOutput.Value
            Case 0  '// 原文の下
                appendIdx = Len(tCell.Text) + 2 + Len(SEPARATOR)    '// 改行コードとセパレータ(--)の分を開始IndexとするSEPARATORは改行コード含。
                tCell.Value = tCell.Text & vbLf & SEPARATOR & resultText
                tCell.Characters(Start:=appendIdx, Length:=Len(resultText)).Font.Color = RGB(0, 0, 255)
            Case 1  '// 原文の上
                appendIdx = Len(tCell.Text)
                tCell.Value = resultText & vbLf & SEPARATOR & vbLf & tCell.Text
                tCell.Characters(Start:=1, Length:=Len(resultText)).Font.Color = RGB(0, 0, 255)
            Case 2  '// 原文を上書き
                tCell.Value = resultText
            Case 3  '// 右のセルに上書き
                tCell.Offset(, 1).Value = resultText
            Case 4  '// コメントに設定
                Call tCell.NoteText(resultText)
                tCell.Comment.Shape.TextFrame.AutoSize = True
                '// コメント幅調整
                If tCell.Comment.Shape.Width > 500 Then
                    tCell.Comment.Shape.Height = tCell.Comment.Shape.Height * (1 / (500 / tCell.Comment.Shape.Width)) * 1.2
                    tCell.Comment.Shape.Width = 500
                End If
        End Select
    Next
    
'    httpReq.Close
    Set httpReq = Nothing
    
    Exit Sub
ErrorHandler:
    Call gsResumeAppEvents
    Call gsShowErrorMsgDlg_VBA("frmTranslate.psTranslate_DeepL", Err)
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// メソッド：   URLエンコード
'// 説明：       引数の文字列をエンコードする。
'//              Excel2013以前でも実行可能にするためWorksheetFunction.EncodeURLは使用しない
'// ////////////////////////////////////////////////////////////////////////////
Private Function EncodeURL2(targetTxt As String, htmlDoc As HTMLDocument) As String
    Dim elm         As HTMLHtmlElement
    
    targetTxt = Replace(targetTxt, "\", "\\")
'    targetTxt = Replace(targetTxt, "#", "\#")
    targetTxt = Replace(targetTxt, "'", "\'")
    targetTxt = Replace(targetTxt, vbLf, TAG_LINE_FEED)
    'Set htmlDoc = CreateObject("htmlfile")
    
    Set elm = htmlDoc.createElement("span")
    Call elm.setAttribute("id", "result")
    Call htmlDoc.body.appendChild(elm)
    Call htmlDoc.parentWindow.execScript("document.getElementById('result').innerText = encodeURIComponent('" & targetTxt & "');", "JScript")
    EncodeURL2 = elm.innerText
    Call elm.ParentNode.RemoveChild(elm)
    Set elm = Nothing
End Function


'// ////////////////////////////////////////////////////////////////////////////
'// メソッド：   結果文字列補正
'// 説明：       翻訳サービスから戻された文字列のタグなどを補正する
'// ////////////////////////////////////////////////////////////////////////////
Private Function DecodeResultText(targetTxt As String) As String
    DecodeResultText = Replace(targetTxt, TAG_LINE_FEED, vbLf)
End Function


'// ////////////////////////////////////////////////////////////////////////////
'// メソッド：   サービス使用状況取得
'// 説明：       翻訳サービスからモニタリング状況を取得し、文字列で戻す
'// ////////////////////////////////////////////////////////////////////////////
Private Function pfGetServiceUsage() As String
On Error GoTo ErrorHandler
    Dim httpReq         As New XMLHTTP60
    Dim reqParam        As String
    Dim sourceText      As String
    Dim resultText      As String
    
    Call httpReq.Open("POST", "https://api-free.deepl.com/v2/usage?")
    Call httpReq.setRequestHeader("Content-Type", "application/x-www-form-urlencoded")
    reqParam = "&auth_key=" & auth_key
    Call httpReq.send(reqParam)
    
    Do While httpReq.readyState < 4
        DoEvents
    Loop
        
    '// レスポンス取得後（VBA-JSONは使用せず、文字列として処理する）
    If httpReq.Status = 200 Then
        pfGetServiceUsage = pfFormatUsage(httpReq.responseText)
    Else
        pfGetServiceUsage = BLANK
    End If
    
'    httpReq.Close
    Set httpReq = Nothing
    Exit Function
    
ErrorHandler:
    Call gsShowErrorMsgDlg_VBA("frmTranslate.pfGetServiceUsage", Err)
End Function


'// ////////////////////////////////////////////////////////////////////////////
'// メソッド：   DeepL使用状況JSONフォーマット
'// 説明：       引数のJSON文字列を、 999 / 999 の形式の文字列で戻す。
'//              pfGetServiceUsageのサブプロシージャ
'//              正規表現で数値以外をすべて取り除き､途中のカンマをスラッシュに置き換える
'// 引数：       str: JSON文字列
'// ////////////////////////////////////////////////////////////////////////////
Private Function pfFormatUsage(str As String) As String
    Dim reg     As Object
    
    Set reg = CreateObject("VBScript.RegExp")
    reg.Pattern = "[^0-9,]"
    reg.Global = True
    pfFormatUsage = reg.Replace(StrConv(str, vbNarrow), "")
    pfFormatUsage = Replace(pfFormatUsage, ",", " / ")
End Function


'// ////////////////////////////////////////////////////////////////////////////
'// END
'// ////////////////////////////////////////////////////////////////////////////

