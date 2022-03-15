VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cStatusGauge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'// ////////////////////////////////////////////////////////////////////////////
'// プロジェクト   : エクセル拡張ツール
'// タイトル       : ステータスバー テキストゲージ表示クラス
'//
'// モジュール     : cStatusGauge
'// 最終修正日     : 2009/10/09
'// 最終修正者     : こいちろ
'//
'// 説明           : ステータスバーにテキストでプログレスバーを表現する
'// ////////////////////////////////////////////////////////////////////////////
'// Copyright (c) by Koichiro.  All Rights Reserved.
'// ////////////////////////////////////////////////////////////////////////////
Option Explicit
Option Base 0


'// ////////////////////////////////////////////////////////////////////////////
'// プライベート変数
Private pMaxVal                       As Double       '// プロパティ値保持：最大値
Private pCurrentVal                   As Double       '// プロパティ値保持：現在値
Private pPercentage                   As Integer      '// プロパティ値保持：パーセント


'///////////////////////////////////////////////////////////////////////////////
'// コンストラクタ
Private Sub Class_Initialize()
  Application.StatusBar = False
  pMaxVal = 100
  pCurrentVal = 0
  pPercentage = 0
End Sub


'///////////////////////////////////////////////////////////////////////////////
'// デストラクタ
Private Sub Class_Terminate()
  Application.StatusBar = False
End Sub


'///////////////////////////////////////////////////////////////////////////////
'// プロパティ: 最大値
Public Property Let MaxVal(newValue As Double)
  pMaxVal = newValue
  Call psPutStatusGauge
End Property

Public Property Get MaxVal() As Double
  MaxVal = pMaxVal
End Property


'///////////////////////////////////////////////////////////////////////////////
'// プロパティ: 現在値
Public Property Let CurrentVal(newValue As Double)
  pCurrentVal = newValue
  Call psPutStatusGauge
End Property

Public Property Get CurrentVal() As Double
  CurrentVal = pMaxVal
End Property


'///////////////////////////////////////////////////////////////////////////////
'// プロパティ: パーセンテージ (ReadOnly)
Public Property Get Percentage() As Integer
  Percentage = pPercentage
End Property


'// ////////////////////////////////////////////////////////////////////////////
'// メソッド：   値の増加
'// 説明：       引数の値を増分として受け取り、ゲージに反映させる。
'// 引数：       val2add: 増分
'///////////////////////////////////////////////////////////////////////////////
Public Function addValue(value2add As Double)
  pCurrentVal = pCurrentVal + value2add
  Call psPutStatusGauge
End Function


'// ////////////////////////////////////////////////////////////////////////////
'// メソッド：   ゲージ表示
'// 説明：       ステータスバーにゲージを表示する。
'// ////////////////////////////////////////////////////////////////////////////
Private Sub psPutStatusGauge()
  Dim newVal  As Integer
  
  newVal = CInt(((pCurrentVal / pMaxVal) * 100) - 0.5)
  If newVal <> pPercentage Then
    pPercentage = newVal
    Application.StatusBar = CStr(pPercentage) & "%  " & String(CInt((newVal - 0.5) / 2), "|")
  End If
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// END
'// ////////////////////////////////////////////////////////////////////////////

