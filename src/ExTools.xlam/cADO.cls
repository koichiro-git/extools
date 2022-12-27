VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cADO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'// ////////////////////////////////////////////////////////////////////////////
'// プロジェクト   : エクセル拡張ツール
'// タイトル       : DB/Excel接続クラス(ADO)
'// モジュール     : cADO
'// 説明           : ADO接続を提供
'// ////////////////////////////////////////////////////////////////////////////
'// Copyright (c) by Koichiro.
'// ////////////////////////////////////////////////////////////////////////////
Option Explicit
Option Base 0

'// ////////////////////////////////////////////////////////////////////////////
'// 列挙 宣言
'// 接続DB
Public Enum udDbConnType
  dct_none = -1
'  dct_oracle = 0
  dct_odbc = 1
'  dct_db2 = 1
'  dct_sqlsrv = 2
'  dct_mysql = 3
  dct_excel = 2
End Enum


'// ////////////////////////////////////////////////////////////////////////////
'// プライベート変数
Private AdoConnection                   As Object       '// コネクションを参照
Private pNativeError                    As Long         '// プロパティ値保持：エラーコード（DBネイティブ）
Private pErrorCode                      As Long         '// プロパティ値保持：エラーコード
Private pErrorText                      As String       '// プロパティ値保持：エラーメッセージ
Private pDbType                         As udDbConnType '// プロパティ値保持：データベース種類
Private pDataSourceName                 As String       '// プロパティ値保持：データソース名
Private pDmlRows                        As Long         '// プロパティ値保持：DML更新行数

'///////////////////////////////////////////////////////////////////////////////
'// コンストラクタ
Private Sub Class_Initialize()
  Set AdoConnection = Nothing
  Call InitError
End Sub


'///////////////////////////////////////////////////////////////////////////////
'// デストラクタ
Private Sub Class_Terminate()
  If Not AdoConnection Is Nothing Then
    AdoConnection.Close
  End If
  
  Set AdoConnection = Nothing
  Call InitError
End Sub


'///////////////////////////////////////////////////////////////////////////////
'// プロパティ: エラーコード (ReadOnly)
Public Property Get ErrorCode() As Long
  ErrorCode = pErrorCode
End Property


'///////////////////////////////////////////////////////////////////////////////
'// プロパティ: エラーコード（DBネイティブ） (ReadOnly)
Public Property Get NativeError() As Long
  NativeError = pNativeError
End Property


'///////////////////////////////////////////////////////////////////////////////
'// プロパティ: エラーメッセージ (ReadOnly)
Public Property Get ErrorText() As String
  ErrorText = pErrorText
End Property


'///////////////////////////////////////////////////////////////////////////////
'// プロパティ: DB種類 (ReadOnly)
Public Property Get DbType() As udDbConnType
  DbType = pDbType
End Property

'///////////////////////////////////////////////////////////////////////////////
'// プロパティ: データソース名 (ReadOnly)
Public Property Get DataSourceName() As String
  DataSourceName = pDataSourceName
End Property

'///////////////////////////////////////////////////////////////////////////////
'// プロパティ: DML更新行数 (ReadOnly)
Public Property Get DmlRows() As Long
  DmlRows = pDmlRows
End Property


'// ////////////////////////////////////////////////////////////////////////////
'// メソッド：   エラー初期化
'// 説明：       エラー保持変数を初期化する
'// 引数：       なし
'// 戻り値：     なし
'// 修正履歴：   なし
'// ////////////////////////////////////////////////////////////////////////////
Public Sub InitError()
  pErrorCode = 0
  pErrorText = Space(0)
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// メソッド：   エラー設定
'// 説明：       エラー保持変数にエラーを設定する
'// 引数：       なし
'// 戻り値：     なし
'// 修正履歴：   なし
'// ////////////////////////////////////////////////////////////////////////////
Private Sub SetError()
  If AdoConnection.Errors.Count > 0 Then
    pNativeError = AdoConnection.Errors(0).NativeError
    pErrorCode = AdoConnection.Errors(0).Number
    pErrorText = AdoConnection.Errors(0).Description
    Call AdoConnection.Errors.Clear
  Else
    pNativeError = 0
    pErrorCode = Err.Number
    pErrorText = Err.Description
  End If
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// メソッド：   接続初期化 (ADO)
'// 説明：       接続を初期化する
'// 引数：       HostString: 接続文字列 / 接続先エクセルファイル名
'//              userID:     ユーザID
'//              password:   パスワード
'//              connType:   接続先DB種類
'// 戻り値：     処理の成否  True:成功 False:失敗
'// 修正履歴：   V2より、DML実行のためReadOnlyを0（更新可能）に変更
'// ////////////////////////////////////////////////////////////////////////////
Public Function Initialize(HostString As String, userID As String, password As String, connType As udDbConnType) As Boolean
'On Error GoTo ErrorHandler
    Dim strCNString As String
    
    Call InitError
    Set AdoConnection = CreateObject("ADODB.Connection")
    
    AdoConnection.CursorLocation = adUseClient    '// DML実行時に更新行数を得るための設定
    
    Select Case connType
'    Case dct_oracle
'      Call AdoConnection.Open("Provider=MSDAORA;Data Source=" & HostString & ";User ID=" & userID & ";Password=" & password & ";ReadOnly=0")
        Case dct_odbc
            Call AdoConnection.Open("Provider=MSDASQL;DSN=" & HostString & ";User ID=" & userID & ";Password=" & password & ";ReadOnly=0")
        Case dct_excel
'            Call AdoConnection.Open("Driver={Microsoft Excel Driver (*.xls)};DBQ=" & HostString & ";Extended Properies=""Excel 12.0"";HDR=1;ReadOnly=0")
            Call AdoConnection.Open("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & HostString & ";Extended Properties=""Excel 12.0;HDR=1;ReadOnly=0"";")
'strCNString = "Provider=Microsoft.ACE.OLEDB.12.0;" _
& "Data Source=" & HostString & ";" _
& "Extended Properties=""Excel 12.0;" _
& "HDR=Yes"";"

'Call AdoConnection.Open(strCNString)
    
    End Select
  
    pDbType = connType
    pDataSourceName = HostString
    Initialize = True
    Exit Function

ErrorHandler:
    Call SetError
    Initialize = False
    Set AdoConnection = Nothing
    pDbType = dct_none
End Function


'// ////////////////////////////////////////////////////////////////////////////
'// メソッド：   SQL実行(SELECT、DML)
'// 説明：       SQL文を実行する
'// 引数：       sql: SQL文
'// 戻り値：     レコードセット
'// 修正履歴：   なし
'// ////////////////////////////////////////////////////////////////////////////
Public Function GetRecordset(sql As String) As Object
On Error GoTo ErrorHandler
    Dim cmd As Object
'    Dim rst As Object
  
    Set cmd = CreateObject("ADODB.Command")
    Set cmd.ActiveConnection = AdoConnection
    cmd.CommandType = adCmdText
    cmd.CommandText = sql
    Set GetRecordset = cmd.Execute(pDmlRows)

  
'  Set rst = CreateObject("ADODB.RecordSet")
'  Call rst.Execute(sql, AdoConnection, 0)  '// 0:adOpenForwardOnly
''  Call rst.Open(sql, AdoConnection, 0)  '// 0:adOpenForwardOnly
'  Set GetRecordset = rst
Exit Function
ErrorHandler:
    Call SetError
    Set GetRecordset = Nothing
End Function


'// ////////////////////////////////////////////////////////////////////////////
'// メソッド：   接続クローズ
'// 説明：       ADOセッションを閉じる
'// ////////////////////////////////////////////////////////////////////////////
Public Sub CloseConnection()
    If Not AdoConnection Is Nothing Then
        AdoConnection.Close
    End If
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// END
'// ////////////////////////////////////////////////////////////////////////////

