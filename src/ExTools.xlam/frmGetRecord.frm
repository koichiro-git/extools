VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmGetRecord 
   Caption         =   "SQL�����s"
   ClientHeight    =   9465
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8175
   OleObjectBlob   =   "frmGetRecord.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "frmGetRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'// ////////////////////////////////////////////////////////////////////////////
'// �v���W�F�N�g   : �g���c�[��
'// �^�C�g��       : SQL���s
'// ���W���[��     : frmGetRecord
'// ����           : SELECT �X�N���v�g�̌��ʂ��G�N�Z���ɏo�͂���B
'// ////////////////////////////////////////////////////////////////////////////
'// Copyright (c) by Koichiro.
'// ////////////////////////////////////////////////////////////////////////////
Option Explicit
Option Base 0


'// ////////////////////////////////////////////////////////////////////////////
'// �v���C�x�[�g�ϐ�
Private pFileName           As String   '// �t�@�C����


'// //////////////////////////////////////////////////////////////////
'// �C�x���g�F �������s�{�^�� �N���b�N��
Private Sub cmdExecute_Click()
    Call psExecSearch
End Sub


'// //////////////////////////////////////////////////////////////////
'// �C�x���g�F ���O�C���{�^�� �N���b�N��
Private Sub cmdLogin_Click()
    frmLogin.Show
End Sub


'// //////////////////////////////////////////////////////////////////
'// �C�x���g�F ����{�^�� �N���b�N��
Private Sub cmdClose_Click()
    Call Me.Hide
End Sub


'// //////////////////////////////////////////////////////////////////
'// �C�x���g�F �t�H�[�� ��������
Private Sub UserForm_Initialize()
    Call gsSetCombo(cmbDateFormat, "0,yyyy/mm/dd;1,yyyy/mm/dd hh:mm:ss", 0)
    Call gsSetCombo(cmdHeader, CMB_GRC_HEADER, 0)

    '// �L���v�V�����ݒ�
    frmGetRecord.Caption = LBL_GRC_FORM
    fraOptions.Caption = LBL_GRC_OPTIONS
    cmdLogin.Caption = LBL_GRC_LOGIN
    cmdExecute.Caption = LBL_GRC_SEARCH
    cmdClose.Caption = LBL_COM_CLOSE
    lblDateFormat.Caption = LBL_GRC_DATE_FORMAT
    lblHeader.Caption = LBL_GRC_HEADER
    lblStatement.Caption = LBL_GRC_SCRIPT
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// ���\�b�h�F    �N�G���[���s
'// �����F        �����̃N�G���[�����s���A�V�[�g�ɏo�͂��܂��B
'// ////////////////////////////////////////////////////////////////////////////
Private Sub psExecSearch()
On Error GoTo ErrorHandler
    Dim wkSheet       As Worksheet
    Dim rst           As Object
    Dim headerRows    As Integer
  
    If gADO Is Nothing Then
        Call frmLogin.Show
        If gADO Is Nothing Then
            Exit Sub
        End If
    End If
  
    '// ���C���r�p�k�̖₢���킹
    Application.StatusBar = MSG_QUERY
    Set rst = gADO.GetRecordset(txtScript.Text)
  
    If rst Is Nothing Then
        Call gsShowErrorMsgDlg("frmGetRecord.psExecSearch", Err, gADO)
        Application.StatusBar = False
        Exit Sub
    End If
  
    If rst.Fields.Count > 0 Then    '// SELECT���̏ꍇ
        If Not rst.EOF Then
            Application.ScreenUpdating = False
            
            '// ���[�N�V�[�g��ǉ��B�V�[�g���̓G�N�Z��������
            Set wkSheet = ActiveWorkbook.Worksheets.Add(Count:=1)
            '// ���ʕ\��
            headerRows = pfDrawHeader(wkSheet, rst)    '// �w�b�_�s
            Call psDrawDataRows(wkSheet, rst, headerRows)  ', cmbGroup.Value)   '// �f�[�^�s
            
            '// �y�[�W�ݒ�
            Application.StatusBar = MSG_PAGE_SETUP
            Call gsPageSetup_Lines(wkSheet, headerRows)
            
            '// �R�����g�ݒ�
            Call Selection.NoteText("-- " & Format(Now, "yyyy/mm/dd hh:nn:ss") & vbCrLf & txtScript.Text)
            
            '// �x���\��
            If rst.Fields.Count > Columns.Count Then
              Call MsgBox(MSG_TOO_MANY_COLS, vbOKOnly, APP_TITLE)
            End If
            
            '// �����̐ݒ�
            '//�t�H���g
            wkSheet.Cells.Font.Name = APP_FONT
            wkSheet.Cells.Font.Size = APP_FONT_SIZE
            
            Call wkSheet.Cells(1, 1).Select
        Else
            Call MsgBox(MSG_NO_RESULT, vbOKOnly, APP_TITLE)
        End If
    Else    '// DML�̏ꍇ
        Call MsgBox(gADO.DmlRows & MSG_ROWS_PROCESSED, vbOKOnly, APP_TITLE)
    End If
    
  
    '// �㏈��
    If rst.State = adStateOpen Then
        Call rst.Close
    End If
    
    Set rst = Nothing
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Call Me.Hide
    Exit Sub
  
ErrorHandler:
    Call gsShowErrorMsgDlg("frmGetRecord.psExecSearch", Err, gADO)
    Application.StatusBar = False
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// ���\�b�h�F    ��w�b�_�`��
'// �����F        ��w�b�_��`�悵�܂��B
'// �����F        wkSheet: ���[�N�V�[�g
'//               rst: ���R�[�h�Z�b�g
'// �߂�l�F      �w�b�_�s��
'// ////////////////////////////////////////////////////////////////////////////
Private Function pfDrawHeader(wkSheet As Worksheet, rst As Object) As Integer
On Error GoTo ErrorHandler
    Dim idx       As Integer
    Dim colStr    As String
    Dim strFormat As String
  
    '// �w�b�_�`��s���i�߂�l�j��ݒ�
    Select Case cmdHeader.Value
        Case 0
            pfDrawHeader = 1
        Case 1
            pfDrawHeader = 3
        Case 2
            pfDrawHeader = 0
    End Select
  
    '// �w�b�_�s�̍���
    For idx = 1 To IIf(rst.Fields.Count > Columns.Count, Columns.Count, rst.Fields.Count)
        With rst.Fields(idx - 1)
            '// �����ݒ� //////////
            Select Case CLng(.Type)
                '// 2:adSmallInt, 3:adInteger, 4:adSingle, 5:adDouble, 6:adCurrency, 16:adTinyInt, 17:adUnsignedTinyInt, 18:adUnsignedSmallInt, 19:adUnsignedInt, 20: adBigInt, 21:adUnsignedBigInt, 131:adNumeric, 139:adVarNumeric
                Case 2, 3, 4, 5, 6, 16, 17, 18, 19, 20, 21, 131, 139
                    strFormat = BLANK
                Case 133, 135                     '// adDBDate, adDBTimeStamp
                    strFormat = cmbDateFormat.List(cmbDateFormat.ListIndex, 1)
                Case 134                          '// 134:adDBTime
                    strFormat = "hh:mm:ss"
                Case Else
                    strFormat = "@"
            End Select
            Call wkSheet.Columns(idx).Select
            Selection.NumberFormatLocal = strFormat
            
            '// ���̐ݒ� //////////
            If cmdHeader.Value <> 2 Then
                wkSheet.Cells(1, idx).NumberFormatLocal = "@"
                wkSheet.Cells(1, idx).Value = .Name
            End If
            
            '// ��`�̏o�́i�^�E�����j//////////
            If cmdHeader.Value = 1 Then
                Select Case CLng(.Type)
                    Case 129, 130                     '// adChar, adWChar
                        wkSheet.Cells(2, idx).Value = "CHAR(" & .DefinedSize & ")"
                    Case 200, 202                     '//adVarChar, adVarWChar
                        wkSheet.Cells(2, idx).Value = "VARCHAR(" & .DefinedSize & ")"
                    Case 2, 18                        '// 2:adSmallInt, 18:adUnsignedSmallInt
                        wkSheet.Cells(2, idx).Value = "SMALLINT"
                    Case 3, 19                        '// 3:adInteger, 19:adUnsignedInt
                        wkSheet.Cells(2, idx).Value = "INTEGER"
                    Case 16, 17                       '// 16:adTinyInt 17:adUnsignedTinyInt
                        wkSheet.Cells(2, idx).Value = "TINYINT"
                    Case 20, 21                       '// 20:adBigInt, 21:adUnsignedBigInt
                        wkSheet.Cells(2, idx).Value = "BIGINT"
                    Case 4                            '// 4:adSingle
                        wkSheet.Cells(2, idx).Value = "SINGLE"
                    Case 5                            '// 5:adDouble
                        wkSheet.Cells(2, idx).Value = "DOUBLE"
                    Case 6                            '// 6:adCurrency
                        wkSheet.Cells(2, idx).Value = "CURRENCY"
                    Case 131, 139                     '// 131:adNumeric, 139:adVarNumeric
                        If .Precision = 0 Then
                            wkSheet.Cells(2, idx).Value = "NUMERIC"
                        ElseIf .NumericScale >= 0 Then
                            wkSheet.Cells(2, idx).Value = "NUMERIC(" & .Precision & "," & .NumericScale & ")"
                        Else
                            wkSheet.Cells(2, idx).Value = "NUMERIC(" & .Precision & ")"
                        End If
                    Case 133                          '// 133:adDBDate
                        wkSheet.Cells(2, idx).Value = "DATE"
                    Case 134                          '// 134:adDBTime
                        wkSheet.Cells(2, idx).Value = "TIME"
                    Case 135                          '// adDBTimeStamp
                        wkSheet.Cells(2, idx).Value = "TIMESTAMP"
                    Case 203  '// lob
                        wkSheet.Cells(2, idx).Value = "CLOB"
                    Case Else
                        wkSheet.Cells(2, idx).Value = BLANK
                End Select
                wkSheet.Cells(3, idx).Value = "-"
            End If
        End With
    Next
    '// �g���̐ݒ�
    Call wkSheet.Range(Cells(1, pfDrawHeader + 1), Cells(1, wkSheet.UsedRange.Columns.Count)).Select
    Call gsDrawLine_Header
    
    '// �g�̌Œ��ݒ�
    If pfDrawHeader > 0 Then
        Call wkSheet.Activate
        Call wkSheet.Rows(pfDrawHeader + 1).Select
        ActiveWindow.FreezePanes = True
    End If
    Exit Function

ErrorHandler:
    Call gsShowErrorMsgDlg("frmGetRecord.pfDrawHeader", Err)
End Function


'// ////////////////////////////////////////////////////////////////////////////
'// ���\�b�h�F    ���[�`��
'// �����F        �e�s�̒l��`�悵�܂��B
'// �����F        wksheet: ���[�N�V�[�g
'//               rst: ���R�[�h�Z�b�g
'//               headerRows: �w�b�_�s��
'//               groupIdx: �O���[�v�������(V2�Ŕp�~�j
'// �߂�l�F      �Ȃ�
'// ////////////////////////////////////////////////////////////////////////////
Private Sub psDrawDataRows(wkSheet As Worksheet, rst As Object, headerRows As Integer)  ', groupIdx As Integer)
On Error GoTo ErrorHandler
    Dim idxRow          As Long
    Dim idxCol          As Integer
    Dim cntCol          As Integer
    Dim varResult       As Variant    '// ���ʕێ��z��i��,�s�j��redim�̎d�l�Ή��̂��߁A�s�Ɨ��ʏ�Ɣ��΂Ɏ��̂Œ���
    
    idxRow = 0
  
    Do While Not rst.EOF
        '// Variant�z�񐮔�
        If idxRow = 0 Then
            cntCol = rst.Fields.Count
            ReDim varResult(cntCol - 1, 1)
            
        Else
            ReDim Preserve varResult(cntCol - 1, idxRow + 1)
        End If
        idxRow = idxRow + 1
        
        '// �f�[�^��z��i��, �s�j�Ɋi�[
        For idxCol = 0 To IIf(cntCol > Columns.Count - 1, Columns.Count - 1, cntCol - 1)
            varResult(idxCol, idxRow - 1) = IIf(IsNull(rst.Fields(idxCol).Value), BLANK, rst.Fields(idxCol).Value)
'            If (idxCol > groupIdx) Then
'                '// �ŏ����x���ȍ~�̏ꍇ�A�l��`��
'                varResult(idxResult, idxCol) = IIf(IsNull(rst.Fields(idxCol - 1).Value), BLANK, rst.Fields(idxCol - 1).Value)
'            ElseIf (aryLastVal(idxCol - 1) = BLANK) Or (aryLastVal(idxCol - 1) <> rst.Fields(idxCol - 1).Value) Then
'                '// ���O�̒l���قȂ�ꍇ (�� ���O�̒l���u�����N�̏ꍇ)
'                '// �z���̃��x���̒��O�̒l���N���A
'                For aryIdx = groupIdx To idxCol Step -1
'                    aryLastVal(aryIdx - 1) = BLANK
'                Next
'                varResult(idxResult, idxCol) = IIf(IsNull(rst.Fields(idxCol - 1).Value), BLANK, rst.Fields(idxCol - 1).Value)
'                aryLastVal(idxCol - 1) = IIf(IsNull(rst.Fields(idxCol - 1).Value), BLANK, rst.Fields(idxCol - 1).Value)
'            End If
        Next
        Call rst.MoveNext
    Loop
    
    '// Variant�̓��e���s�����ւ��ăV�[�g�ɒ���t��
    wkSheet.Range(wkSheet.Cells(headerRows + 1, 1), wkSheet.Cells(idxRow + headerRows, cntCol)).Value = WorksheetFunction.Transpose(varResult)
    
    '// �r����`��
    Call wkSheet.UsedRange.Select
    Call gsDrawLine_Data
    
    Exit Sub

ErrorHandler:
    Call gsShowErrorMsgDlg("frmGetRecord.psDrawDataRows", Err)
End Sub

'// ////////////////////////////////////////////////////////////////////////////
'// END.
'// ////////////////////////////////////////////////////////////////////////////
