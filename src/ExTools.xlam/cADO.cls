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
'// �v���W�F�N�g   : �G�N�Z���g���c�[��
'// �^�C�g��       : DB/Excel�ڑ��N���X(ADO)
'//
'// ���W���[��     : cADO
'// �ŏI�C����     : 2009/10/14 00:00:00
'// �ŏI�C����     : ��������
'//
'// ����           : ADO�ڑ����
'// ////////////////////////////////////////////////////////////////////////////
'// Copyright (c) by Koichiro.  All Rights Reserved.
'// ////////////////////////////////////////////////////////////////////////////
Option Explicit
Option Base 0

'// ////////////////////////////////////////////////////////////////////////////
'// �� �錾
'// �ڑ�DB
Public Enum udDbConnType
  dct_none = -1
  dct_oracle = 0
  dct_odbc = 1
'  dct_db2 = 1
'  dct_sqlsrv = 2
'  dct_mysql = 3
  dct_excel = 2
End Enum


'// ////////////////////////////////////////////////////////////////////////////
'// �v���C�x�[�g�ϐ�
Private AdoConnection                   As Object       '// �R�l�N�V�������Q��
Private pNativeError                    As Long         '// �v���p�e�B�l�ێ��F�G���[�R�[�h�iDB�l�C�e�B�u�j
Private pErrorCode                      As Long         '// �v���p�e�B�l�ێ��F�G���[�R�[�h
Private pErrorText                      As String       '// �v���p�e�B�l�ێ��F�G���[���b�Z�[�W
Private pDbType                         As udDbConnType '// �v���p�e�B�l�ێ��F�f�[�^�x�[�X���
Private pDmlRows                        As Long         '// �v���p�e�B�l�ێ��FDML�X�V�s��


'///////////////////////////////////////////////////////////////////////////////
'// �R���X�g���N�^
Private Sub Class_Initialize()
  Set AdoConnection = Nothing
  Call InitError
End Sub


'///////////////////////////////////////////////////////////////////////////////
'// �f�X�g���N�^
Private Sub Class_Terminate()
  If Not AdoConnection Is Nothing Then
    AdoConnection.Close
  End If
  
  Set AdoConnection = Nothing
  Call InitError
End Sub


'///////////////////////////////////////////////////////////////////////////////
'// �v���p�e�B: �G���[�R�[�h (ReadOnly)
Public Property Get ErrorCode() As Long
  ErrorCode = pErrorCode
End Property


'///////////////////////////////////////////////////////////////////////////////
'// �v���p�e�B: �G���[�R�[�h�iDB�l�C�e�B�u�j (ReadOnly)
Public Property Get NativeError() As Long
  NativeError = pNativeError
End Property


'///////////////////////////////////////////////////////////////////////////////
'// �v���p�e�B: �G���[���b�Z�[�W (ReadOnly)
Public Property Get ErrorText() As String
  ErrorText = pErrorText
End Property


'///////////////////////////////////////////////////////////////////////////////
'// �v���p�e�B: DB��� (ReadOnly)
Public Property Get DbType() As udDbConnType
  DbType = pDbType
End Property

'///////////////////////////////////////////////////////////////////////////////
'// �v���p�e�B: DML�X�V�s�� (ReadOnly)
Public Property Get DmlRows() As Long
  DmlRows = pDmlRows
End Property


'// ////////////////////////////////////////////////////////////////////////////
'// ���\�b�h�F   �G���[������
'// �����F       �G���[�ێ��ϐ�������������
'// �����F       �Ȃ�
'// �߂�l�F     �Ȃ�
'// �C�������F   �Ȃ�
'// ////////////////////////////////////////////////////////////////////////////
Public Sub InitError()
  pErrorCode = 0
  pErrorText = Space(0)
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// ���\�b�h�F   �G���[�ݒ�
'// �����F       �G���[�ێ��ϐ��ɃG���[��ݒ肷��
'// �����F       �Ȃ�
'// �߂�l�F     �Ȃ�
'// �C�������F   �Ȃ�
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
'// ���\�b�h�F   �ڑ������� (ADO)
'// �����F       �ڑ�������������
'// �����F       HostString: �ڑ������� / �ڑ���G�N�Z���t�@�C����
'//              userID:     ���[�UID
'//              password:   �p�X���[�h
'//              connType:   �ڑ���DB���
'// �߂�l�F     �����̐���  True:���� False:���s
'// �C�������F   V2���ADML���s�̂���ReadOnly��0�i�X�V�\�j�ɕύX
'// ////////////////////////////////////////////////////////////////////////////
Public Function Initialize(HostString As String, userID As String, password As String, connType As udDbConnType) As Boolean
On Error GoTo ErrorHandler
  Call InitError
  Set AdoConnection = CreateObject("ADODB.Connection")
  
  AdoConnection.CursorLocation = adUseClient    '// DML���s���ɍX�V�s���𓾂邽�߂̐ݒ�
  
  Select Case connType
    Case dct_oracle
      Call AdoConnection.Open("Provider=MSDAORA;Data Source=" & HostString & ";User ID=" & userID & ";Password=" & password & ";ReadOnly=0")
    Case dct_odbc
      Call AdoConnection.Open("Provider=MSDASQL;DSN=" & HostString & ";User ID=" & userID & ";Password=" & password & ";ReadOnly=0")
    Case dct_excel
      Call AdoConnection.Open("Driver={Microsoft Excel Driver (*.xls)};DBQ=" & HostString & ";Extended Properies=""Excel 8.0"";HDR=1;ReadOnly=0")
  End Select
  
  pDbType = connType
  Initialize = True
  Exit Function

ErrorHandler:
  Call SetError
  Initialize = False
  Set AdoConnection = Nothing
  pDbType = dct_none
End Function


'// ////////////////////////////////////////////////////////////////////////////
'// ���\�b�h�F   SQL���s(SELECT�ADML)
'// �����F       SQL�������s����
'// �����F       sql: SQL��
'// �߂�l�F     ���R�[�h�Z�b�g
'// �C�������F   �Ȃ�
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
'// END
'// ////////////////////////////////////////////////////////////////////////////
