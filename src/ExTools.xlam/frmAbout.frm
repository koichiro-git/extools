VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAbout 
   Caption         =   "エクセル拡張ツール"
   ClientHeight    =   2805
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6030
   OleObjectBlob   =   "frmAbout.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'// ////////////////////////////////////////////////////////////////////////////
'// プロジェクト   : 拡張ツール
'// タイトル       : About画面フォーム
'// モジュール     : frmAbout
'// 説明           : バージョン情報、リポジトリ情報を出力する
'// ////////////////////////////////////////////////////////////////////////////
'// Copyright (c) by Koichiro.
'// ////////////////////////////////////////////////////////////////////////////
Option Explicit
Option Base 0


'// //////////////////////////////////////////////////////////////////
'// イベント： フォーム 初期化時
Private Sub UserForm_Initialize()
    Me.Caption = APP_TITLE
    lblVersion.Caption = APP_TITLE & Space(1) & APP_VERSION
    lblRepo.Caption = LBL_ABT_REPO
    lblManual.Caption = LBL_ABT_MANUAL
    cmdClose.Caption = LBL_COM_CLOSE
End Sub


'// //////////////////////////////////////////////////////////////////
'// イベント： 閉じるボタン クリック時
Private Sub cmdClose_Click()
  Call Me.Hide
End Sub


'// //////////////////////////////////////////////////////////////////
'// イベント： ラベル_ダウンロード クリック時
Private Sub lblDownload_Click()
    Call ThisWorkbook.FollowHyperlink(Address:="https://github.com/koichiro-git/extools/releases")
End Sub


'// //////////////////////////////////////////////////////////////////
'// イベント： ラベル_マニュアル クリック時
Private Sub lblUserManual_Click()
    Call ThisWorkbook.FollowHyperlink(Address:="https://koichiro-git.github.io/extools/")
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// END
'// ////////////////////////////////////////////////////////////////////////////
