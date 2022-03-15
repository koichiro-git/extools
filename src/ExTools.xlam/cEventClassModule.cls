VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cEventClassModule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'// ////////////////////////////////////////////////////////////////////////////
'// プロジェクト   : 拡張ツール
'// タイトル       : アプリケーション・イベント定義
'//
'// モジュール     : cEventClassModule
'// 最終修正日     : 2022/01/08
'// 最終修正者     : こいちろ
'//
'// 説明           : Excel2007以前にメニューバーの有効／無効を切り替えるためのAPイベントを
'//                  定義していた。
'//                  リボン採用でこのイベントが不要となったこと、およびSQL実行機能で
'//                  Excel以外のデータソースを廃止したことに伴い、この定義も廃止する。
'//                  将来、このモジュール自体を削除する。
'// ////////////////////////////////////////////////////////////////////////////
'// Copyright (c) by Koichiro.
'// ////////////////////////////////////////////////////////////////////////////
Option Explicit

Public WithEvents App As Application
Attribute App.VB_VarHelpID = -1



'// ////////////////////////////////////////////////////////////////////////////
'// END
'// ////////////////////////////////////////////////////////////////////////////
