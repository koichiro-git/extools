Attribute VB_Name = "mdlModPhoneNumber"
'// ////////////////////////////////////////////////////////////////////////////
'// プロジェクト   : 拡張ツール 追加パック
'// タイトル       : 電話番号ハイフン区切り
'// モジュール     : mdlModPhoneNumber
'// 説明           : システムの共通関数、起動時の設定などを管理
'//                  国内電話番号のみ対応。国番号（+81等）は対応していない
'// ////////////////////////////////////////////////////////////////////////////
'// Copyright (c) by Koichiro.
'// ////////////////////////////////////////////////////////////////////////////
Option Explicit
Option Base 0

'// ////////////////////////////////////////////////////////////////////////////
'// プライベート定数

'// 市外局番リスト(元ネタ＝総務省 https://www.soumu.go.jp/main_sosiki/joho_tsusin/top/tel_number/shigai_list.html)
Private Const AREA_CODE_LIST = _
    "011/0123/0124/0125/0126/01267/0133/0134/0135/0136/01372/01374/0137/01377/0138/01392/0139/01397/01398/0142/0143/0144/0145/01456/01457/0146/01466/0152/0153/0154/01547/015/0155/01558/0156/01564/0157/0158/01586/01587/0162/01632/01634/01635/0163/0164/01648/0165/01654/01655/01656/01658/0166/0167/0172/0173/0174/0175/0176/017/0178/0179/0182/0183/0184/0185/0186/0187/018/0191/0192/0193/0194/0195/019/0197/0198/" & _
    "022/0220/0223/0224/0225/0226/0228/0229/0233/0234/0235/023/0237/0238/0240/0241/0242/0243/0244/024/0246/0247/0248/025/0250/0254/0255/0256/0257/0258/0259/0260/0261/026/0263/0264/0265/0266/0267/0268/0269/0270/027/0274/0276/0277/0278/0279/0280/0282/0283/0284/0285/028/0287/0288/0289/0291/029/0293/0294/0295/0296/0297/0299/" & _
    "03/0422/042/0428/04/043/0436/0438/0439/044/045/0460/046/0463/0465/0466/0467/0470/047/0475/0476/0478/0479/048/0480/049/0493/0494/0495/04992/04994/04996/04998/" & _
    "052/053/0531/0532/0533/0536/0537/0538/0539/054/0544/0545/0547/0548/0550/0551/055/0553/0554/0555/0556/0557/0558/0561/0562/0563/0564/0565/0566/0567/0568/0569/0572/0573/0574/0575/0576/05769/0577/0578/058/0581/0584/0585/0586/0587/059/0594/0595/0596/0597/05979/0598/0599/" & _
    "06/072/0721/0725/073/0735/0736/0737/0738/0739/0740/0742/0743/0744/0745/0746/07468/0747/0748/0749/075/0761/076/0763/0765/0766/0767/0768/0770/0771/0772/0773/0774/077/0776/0778/0779/078/0790/0791/079/0794/0795/0796/0797/0798/0799/" & _
    "082/0820/0823/0824/0826/0827/0829/083/0833/0834/0835/0836/0837/0838/08387/08388/08396/0845/0846/0847/08477/0848/084/08512/08514/0852/0853/0854/0855/0856/0857/0858/0859/086/0863/0865/0866/0867/0868/0869/0875/0877/087/0879/0880/0883/0884/0885/088/0887/0889/0892/0893/0894/0895/0896/0897/0898/089/" & _
    "092/0920/093/0930/0940/0942/0943/0944/0946/0947/0948/0949/0950/0952/0954/0955/0956/0957/095/0959/096/0964/0965/0966/0967/0968/0969/0972/0973/0974/097/0977/0978/0979/098/0980/09802/0982/0983/0984/0985/0986/0987/09912/09913/099/0993/0994/0995/0996/09969/0997/"

'// 携帯電話・IP Phoneなど11桁局番
Private Const MOBILE_IP_CODE_LIST = "090/080/070/060/050"


'// ////////////////////////////////////////////////////////////////////////////
'// メソッド：   リボンボタンコールバック管理
'// 説明：       リボンからのコールバックをつかさどる
'//              押されたコントロールのIDを基に処理を呼び出す。
'// 引数：       control 対象コントロール
'// ////////////////////////////////////////////////////////////////////////////
Public Sub ribbonCallback_PhoneNum(control As IRibbonControl)
    Select Case control.ID
        Case "FormatPhoneNumbers"                       '// 電話番号補正
            Call psFormatPhoneNumbers
    End Select
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// メソッド：   電話番号ハイフン付与 主関数
'// 説明：       選択範囲の文字列を電話番号と見なしてハイフンを付与する
'// ////////////////////////////////////////////////////////////////////////////
Private Sub psFormatPhoneNumbers()
On Error GoTo ErrorHandler
    Dim tCell       As Range    '// 変換対象セル
    Dim bff         As String   '// 変換後文字列格納バッファ
    Dim cnt         As Integer  '// エラーカウント
    
    '// 事前チェック（アクティブシート保護、選択タイプ＝セル）
    If Not gfPreCheck(protectCont:=True, selType:=TYPE_RANGE) Then
        Exit Sub
    End If
    
    Call gsSuppressAppEvents
    
    cnt = 0
    If Selection.Count > 1 Then
        For Each tCell In Selection.SpecialCells(xlCellTypeConstants, xlNumbers + xlTextValues)  '//SELECTIONが空の場合はエラーハンドラでキャッチ
            bff = pfApplyPhoneNumberFormat(tCell.Text)
            If bff <> BLANK Then                '// 変換ロジックからブランクが戻された場合は無視
                tCell.Value = bff
            ElseIf ActiveCell.Text <> BLANK Then    '// 変換に失敗し、セル値が空白でない場合にはエラー扱いとしてカウント
                cnt = cnt + 1
            End If
        Next
    Else
        bff = pfApplyPhoneNumberFormat(ActiveCell.Text)
        If bff <> BLANK Then                    '// 変換ロジックからブランクが戻された場合は無視
            ActiveCell.Value = pfApplyPhoneNumberFormat(bff)
        End If
    End If
    
    Call gsResumeAppEvents
    
    If cnt > 0 Then
        Call MsgBox(MSG_ERR & "(" & cnt & ")", vbOKOnly, APP_TITLE)
    End If
    
    Exit Sub
ErrorHandler:
    Call gsResumeAppEvents
    If Err.Number = 1004 Then  '// 範囲選択が正しくない場合
        Call MsgBox(MSG_INVALID_RANGE, vbOKOnly, APP_TITLE)
    Else
        Call gsShowErrorMsgDlg_VBA("mdlModPhoneNumber.psFormatPhoneNumbers", Err)
    End If
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// メソッド：   電話番号ハイフン付与
'// 説明：       引数の文字列に市外局番に応じたハイフンを付与する
'//              該当する市外局番が無い場合、無効な文字列の場合はブランクを返す
'//              psFormatPhoneNumbers から呼び出される実処理
'// ////////////////////////////////////////////////////////////////////////////
Private Function pfApplyPhoneNumberFormat(originalStr As String) As String
    Dim targetVal As String
    Dim digs As Integer
    Dim rslt As String
    
    targetVal = pfPicNumerics(originalStr)
    
    If targetVal = BLANK Then '// 空白
        rslt = BLANK
    ElseIf Left(targetVal, 1) <> "0" Then  '// 先頭がゼロでないばあい
        rslt = BLANK
    ElseIf InStr(1, MOBILE_IP_CODE_LIST, Left(targetVal, 3)) > 0 Then
        '// 携帯・IP Phone
        rslt = Left(targetVal, 3) & "-" & Mid(targetVal, 4, 4) & "-" & Mid(targetVal, 8)
        If Len(rslt) <> 13 Then
            rslt = BLANK
        End If
    Else
        '// 固定電話（市外局番の長いものから順にチェック）
        For digs = 5 To 2 Step -1
            If InStr(1, AREA_CODE_LIST, Left(targetVal, digs)) > 0 Then
                rslt = Left(targetVal, digs) & "-" & Mid(targetVal, digs + 1, 6 - digs) & "-" & Mid(targetVal, 7)
                Exit For
            End If
        Next
        If Len(rslt) <> 12 Then
            rslt = BLANK
        End If
    End If
    
    pfApplyPhoneNumberFormat = rslt
End Function


'// ////////////////////////////////////////////////////////////////////////////
'// 引数の文字列から数値のみを抽出し半角で戻す
'// 引数 str 変換対象文字列
'// ////////////////////////////////////////////////////////////////////////////
Private Function pfPicNumerics(str As String) As String
    Dim reg     As Object
    Dim rslt    As String
    
    Set reg = CreateObject("VBScript.RegExp")
    reg.Pattern = "\D"
    reg.Global = True
    pfPicNumerics = reg.Replace(StrConv(str, vbNarrow), "")
End Function


'// ////////////////////////////////////////////////////////////////////////////
'// END
'// ////////////////////////////////////////////////////////////////////////////
