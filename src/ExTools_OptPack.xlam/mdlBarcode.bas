Attribute VB_Name = "mdlBarcode"
'// ////////////////////////////////////////////////////////////////////////////
'// プロジェクト   : 拡張ツール 追加パック
'// タイトル       : QRコード表示
'// モジュール     : mdlBarcode
'// 説明           : セル内容からQRコード画像を生成する。
'//                  Access用ランタイムはExcelでは動作保証されていないため、フリーのWebサービスを使用
'// ////////////////////////////////////////////////////////////////////////////
'// Copyright (c) by Koichiro.
'// ////////////////////////////////////////////////////////////////////////////
Option Explicit
Option Base 0


'// ////////////////////////////////////////////////////////////////////////////
'// メソッド：   リボンボタンコールバック管理
'// 説明：       リボンからのコールバックをつかさどる
'//              押されたコントロールのIDを基に処理を呼び出す。
'// 引数：       control 対象コントロール
'// ////////////////////////////////////////////////////////////////////////////
Public Sub ribbonCallback_BarCode(control As IRibbonControl)
    Select Case control.ID
        Case "QRCode"                       '// QRコード
            Call psDrawBarCode
    End Select
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// メソッド：   バーコード描画
'// 説明：       選択されたセルの値を元にバーコードを描画する
'// ////////////////////////////////////////////////////////////////////////////
Private Sub psDrawBarCode()
    Dim tCell     As Range    '// 変換対象セル
    
    '// 事前チェック（アクティブシート保護、選択タイプ＝セル）
    If Not gfPreCheck(protectCont:=True, selType:=TYPE_RANGE) Then
        Exit Sub
    End If
    
    Call gsSuppressAppEvents
    
    If Selection.Count > 1 Then
        For Each tCell In Selection.SpecialCells(xlCellTypeConstants, xlNumbers + xlTextValues)
            Call psDrawBarCode_sub(tCell)
        Next
    Else
        Call psDrawBarCode_sub(ActiveCell)
    End If
    
    Call gsResumeAppEvents
    Exit Sub
    
ErrorHandler:
    Call gsResumeAppEvents
    Call gsShowErrorMsgDlg_VBA("mdlBarcode.psDrawBarCode", Err)
End Sub

    
'// ////////////////////////////////////////////////////////////////////////////
'// メソッド：   バーコード描画（サブ）
'// 説明：       フリーのQRコード描画サービスにアクセスし、引数のセルにシェイプとして張り付ける
'// 引数：       tCell: シェイプ貼り付け位置
'// ////////////////////////////////////////////////////////////////////////////
Private Sub psDrawBarCode_sub(tCell As Range)
    Dim obj         As Shape    '// 画像貼付用シェイプ
    
    Set obj = ActiveSheet.Shapes.AddPicture("http://api.qrserver.com/v1/create-qr-code/?data=" & tCell.Text & _
                                            "!&size=300x300", False, True, tCell.Left, tCell.Top, tCell.Width, tCell.Height)
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// END
'// ////////////////////////////////////////////////////////////////////////////
