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

'// 古いバージョンでもコンパイルを通すためバーコードの参照設定は行わない。このため定数を改めて定義
'//https://rdr.utopiat.net/」

'Enum BarcodeStyle
'    JAN13 = 2
'    JAN8 = 3
'    Casecode = 4
'    NW7 = 5
'    Code39 = 6
'    Code128 = 7
'    QR = 11
'End Enum


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
Sub psDrawBarCode()
Attribute psDrawBarCode.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim tCell     As Range    '// 変換対象セル
    
    '//
'    If Application.Version < 16 Then
'        Call MsgBox("バーコードはExcel2016以降のバージョンでのみ使用可能です")
'        Exit Sub
'    End If
    
    '// 事前チェック（アクティブシート保護、選択タイプ＝セル）
    If Not gfPreCheck(protectCont:=True, selType:=TYPE_RANGE) Then
        Exit Sub
    End If
    
    Call gsSuppressAppEvents
    
    If Selection.Count > 1 Then
        For Each tCell In Selection.SpecialCells(xlCellTypeConstants, xlNumbers + xlTextValues)
            Call psDrawBarCode_sub2(tCell)
        Next
    Else
        Call psDrawBarCode_sub2(ActiveCell)
    End If
    
    Call gsResumeAppEvents
    Exit Sub
    
ErrorHandler:
    '//
    Call gsResumeAppEvents
End Sub
    
    
Private Sub psDrawBarCode_sub(tCell As Range)
    Dim obj         As OLEObject    '// バーコードのActiveXオブジェクト本体を格納
    Dim bcd         As Object       '// ActiveXオブジェクト内部のバーコード描画オブジェクトを格納
    Dim pctCamera   As Object       '// 画像に変換する際に一時的にCopyPictureに画像として認識させるためのカメラオブジェクト
        
    Set obj = ActiveSheet.OLEObjects.Add(ClassType:="BARCODE.BarCodeCtrl.1", Link:=False, DisplayAsIcon:=False)
    
    '// 設定(プロパティページ分)
    Set bcd = obj.Object
    With bcd
        .style = 11
        .Validation = 1
'                .Refresh
    End With
    
    obj.LinkedCell = tCell.Address      '// Linked Cell をStyleより後に最後に設定しないと、セル値が消える
    obj.Top = tCell.Top                 '// 最後にサイズ変更することで描画リフレッシュ
    obj.Left = tCell.Left
    obj.Width = tCell.Width
    obj.Height = tCell.Height
    
    '// カメラとして複製（CopyPictureではActiveXオブジェクトは認識されないため必要な処理）
    tCell.Select
    tCell.Copy
    Set pctCamera = ActiveSheet.Pictures.Paste(Link:=True)
    '// セルを画像としてコピー・貼り付けする。
    Call tCell.CopyPicture(xlPrinter, xlPicture)
    Call ActiveSheet.Paste
    '// カメラを削除
    pctCamera.Delete
    '// ActiveXオブジェクトを削除
    obj.Delete
End Sub

    
    
Private Sub psDrawBarCode_sub2(tCell As Range)
    Dim obj         As Shape    '// 画像貼付用シェイプ
    
    Set obj = ActiveSheet.Shapes.AddPicture("http://api.qrserver.com/v1/create-qr-code/?data=" & tCell.Text & _
                                            "!&size=300x300", False, True, tCell.Left, tCell.Top, tCell.Width, tCell.Height)
'    obj.Top = tCell.Top
'    obj.Left = tCell.Left
'    obj.Width = tCell.Width
'    obj.Height = tCell.Height
    
End Sub
    
    
'// カメラを使ってコピー
'    Selection.Copy
'    Range("M13").Select
'    ActiveSheet.Pictures.Paste(Link:=True).Select
'    ActiveSheet.Shapes.Range(Array("Picture 60")).Select
'    Application.CutCopyMode = False
'    Selection.Copy
'    Range("M14").Select
'    ActiveSheet.Pictures.Paste.Select
    
    
    
    
'
'
'
'
'    For idxRow = 7 To 12
'
'        Set obj = ActiveSheet.OLEObjects.Add(ClassType:="BARCODE.BarCodeCtrl.1", Link:=False, DisplayAsIcon:=False)
'        obj.Top = Cells(idxRow, 2).Top
'        obj.Left = Cells(idxRow, 2).Left
'        obj.Width = Cells(idxRow, 2).Width
'        obj.Height = Cells(idxRow, 2).Height
'
' '       Set bcd = CreateObject("BARCODE.BarCodeCtrl.1")
'        Set bcd = obj.Object
'
'        With bcd
'          .Style = BarcodeStyle.Code128
''          .LinkedCell = Cells(idxRow, 2).Address
''          .SubStyle = BC_Substyle
'          .Validation = 1
''          .LineWeight = BC_LineWeight
''          .Direction = BC_Direction
''          .ShowData = BC_ShowData
''          .ForeColor = BC_ForeColor
''          .BackColor = BC_BackColor
'          .Refresh
'         End With
'        'Linked Cell は最後に設定しないと、セル値が消える。。。
'        obj.LinkedCell = Cells(idxRow, 2).Address
'
'    Next
'
    'JAN,CODE39,ITF,NW-7,CODE128など、
    'JAN/EAN/UPC ITF CODE39 NW-7(CODABAR) CODE128 代表5種類
    



' 'プロパティについては以下URLのMSDN参照
'    'https://msdn.microsoft.com/ja-jp/library/cc427149.aspx
'
'    Const BC_Style As Integer = 7
'    'スタイル
'    '0: UPC-A, 1: UPC-E, 2: JAN-13, 3: JAN-8, 4: Casecode, 5: NW-7,
'    '6: Code-39, 7: Code-128, 8: U.S. Postnet, 9: U.S. Postal FIM, 10: 郵便物の表示用途（日本）
'
'    Const BC_Substyle As Integer = 0
'    'サブスタイル (下記URL参照)
'    'https://msdn.microsoft.com/ja-jp/library/cc427156.aspx
'
'    Const BC_Validation As Integer = 1
'    'データの確認
'    '0: 確認無し, 1: 無効なら計算を補正, 2: 無効なら非表示
'    'Code39/NW-7の場合、「1」でスタート/ストップ文字(*)を自動的に追加
'
'    Const BC_LineWeight As Integer = 3
'    '線の太さ
'    '0: 極細線, 1:細線, 2:中細線, 3:標準, 4:中太線, 5: 太線, 6:極太線, 7:超極太線
'
'    Const BC_Direction As Integer = 0
'    'バーコードの表示方向
'    '0: 0度, 1: 90度, 2: 180度, 3: 270度　[0]が標準
'
'    Const BC_ShowData As Integer = 1
'    'データの表示
'    '0: 表示無し, 1:表示有り
'
'    Const BC_ForeColor As Long = rgbBlack
'    '前景色の指定
'
'    Const BC_BackColor As Long = rgbWhite
'    '背景色の指定
'
'    'rgbBlackなどの色定数は以下URLのMSDN参照
'    'https://msdn.microsoft.com/ja-jp/VBA/Excel-VBA/articles/xlrgbcolor-enumeration-excel
'
'   '**バーコード化の処理＊＊

'End Sub


'// https://translate.google.pl/?sl=en&tl=ja&text=hello&op=translate



