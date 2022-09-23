VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmLogin 
   Caption         =   "ログイン"
   ClientHeight    =   3255
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3510
   OleObjectBlob   =   "frmLogin.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'// ////////////////////////////////////////////////////////////////////////////
'// プロジェクト   : 拡張ツール
'// タイトル       : 認証画面
'// モジュール     : frmLogin
'// 説明           : 接続先への認証を行う。
'// ////////////////////////////////////////////////////////////////////////////
'// Copyright (c) by Koichiro.
'// ////////////////////////////////////////////////////////////////////////////
Option Explicit
Option Base 0

'// ////////////////////////////////////////////////////////////////////////////
'// プライベート定数
'// なし


'// ////////////////////////////////////////////////////////////////////////////
'// イベント： フォーム 初期化時
Private Sub UserForm_Initialize()
    '// コンボボックス設定
    Call gsSetCombo(cmbMethod, dct_excel & ",Excel;" & dct_odbc & ",ODBC", 0)
    
    '// キャプション設定
    frmLogin.Caption = LBL_LGI_FORM
    lblUid.Caption = LBL_LGI_UID
    lblPassword.Caption = LBL_LGI_PASSWORD
    lblConnStr.Caption = LBL_LGI_STRING
    lblConnTo.Caption = LBL_LGI_CONN_TO
    cmdOk.Caption = LBL_LGI_LOGIN
    cmdCancel.Caption = LBL_LGI_CANCEL
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// イベント： フォーム アクティブ時
Private Sub UserForm_Activate()
    If cmbMethod.Value <> dct_excel Then
        '// ユーザIDまたはパスワードにフォーカスを設定
        If txtUserId.Text = BLANK Then
            Call txtUserId.SetFocus
        Else
            Call txtPassword.SetFocus
        End If
    End If
    '// パスワードは常に空
    txtPassword.Text = BLANK
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// イベント： ＯＫボタン クリック時
Private Sub cmdOk_Click()
On Error GoTo ErrorHandler
    Dim isConnected As Boolean
    
    '// 既存のコネクションを切断する
    If Not gADO Is Nothing Then
        Call gADO.CloseConnection
    End If
    Set gADO = Nothing
  
    '// 入力チェック
    If (cmbMethod.Value = dct_excel) Then     '// Excelへの接続では現在のファイルチェック
        If (ActiveWorkbook.Path = BLANK) Or (Not ActiveWorkbook.Saved) Then
            If MsgBox(MSG_NEED_EXCEL_SAVED, vbYesNo, APP_TITLE) = vbNo Then
                cmbMethod.SetFocus
                Exit Sub
            Else
                Call ActiveWorkbook.Save
            End If
        End If
    ElseIf (txtUserId.Text = BLANK) Then      '// ユーザID
        Call MsgBox(MSG_NEED_FILL_ID, vbOKOnly, APP_TITLE)
        Call txtUserId.SetFocus
        Exit Sub
    ElseIf (txtPassword.Text = BLANK) Then    '// パスワード
        Call MsgBox(MSG_NEED_FILL_PWD, vbOKOnly, APP_TITLE)
        Call txtPassword.SetFocus
        Exit Sub
    ElseIf (txtHostString.Text = BLANK) Then  '// 接続文字列
        Call MsgBox(MSG_NEED_FILL_TNS, vbOKOnly, APP_TITLE)
        Call txtHostString.SetFocus
        Exit Sub
    End If
  
    '// 接続
    Set gADO = New cADO
    Select Case cmbMethod.Value
'        Case dct_oracle  '// v2よりオラクルは対象外
'            isConnected = gOra.Initialize(txtHostString.Text, txtUserId.Text, txtPassword.Text, dct_oracle)
        Case dct_odbc
            isConnected = gADO.Initialize(txtHostString.Text, txtUserId.Text, txtPassword.Text, dct_odbc)
        Case dct_excel
            isConnected = gADO.Initialize(ActiveWorkbook.Path & "\" & ActiveWorkbook.Name, BLANK, BLANK, dct_excel)
    End Select
  
    '// 判定
    If Not isConnected Then
        Call MsgBox(MSG_LOG_ON_FAILED & vbLf & vbLf _
                & "Error Number: " & gADO.NativeError & vbLf _
                & "Error Description: " & gADO.ErrorText, vbOKOnly, APP_TITLE)
    Else
        Call Me.Hide
        Call MsgBox(MSG_LOG_ON_SUCCESS, vbOKOnly, APP_TITLE)
    End If
    Exit Sub
  
ErrorHandler:
    Call gsShowErrorMsgDlg("frmLogin.cmdOk_Click", Err)
    Set gADO = Nothing
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// イベント： キャンセルボタン クリック時
Private Sub cmdCancel_Click()
    Call Me.Hide
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// イベント： 接続先コンボ 変更時
Private Sub cmbMethod_Change()
    '// エクセル接続時は、ユーザID、パスワード、接続文字列は使用不可に設定
    If cmbMethod.Value = dct_excel Then
        txtUserId.Enabled = False
        txtUserId.BackColor = CLR_DISABLED
        txtPassword.Enabled = False
        txtPassword.BackColor = CLR_DISABLED
        txtHostString.Enabled = False
        txtHostString.BackColor = CLR_DISABLED
    Else
        txtUserId.Enabled = True
        txtUserId.BackColor = CLR_ENABLED
        txtPassword.Enabled = True
        txtPassword.BackColor = CLR_ENABLED
        txtHostString.Enabled = True
        txtHostString.BackColor = CLR_ENABLED
    End If
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// END
'// ////////////////////////////////////////////////////////////////////////////
