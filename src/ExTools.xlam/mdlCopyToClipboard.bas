Attribute VB_Name = "mdlCopyToClipboard"
'// ////////////////////////////////////////////////////////////////////////////
'// �v���W�F�N�g   : �g���c�[��
'// �^�C�g��       : �N���b�v�{�[�h�փR�s�[�@�\
'// ���W���[��     : mdlCopyToClipboard
'// ����           : �I��͈͂��Œ蒷�A�}�[�N�_�E���`���ŃN���b�v�{�[�h�ɃR�s�[����
'// ////////////////////////////////////////////////////////////////////////////
'// Copyright (c) by Koichiro.
'// ////////////////////////////////////////////////////////////////////////////
Option Explicit
Option Base 0

'// ////////////////////////////////////////////////////////////////////////////
'// �A�v���P�[�V�����萔
Private Const MAX_COL_LEN             As Integer = 80                                               '// �N���b�v�{�[�h�ɃR�s�[����ۂ̗�ő咷


'// ////////////////////////////////////////////////////////////////////////////
'// ���\�b�h�F   ���{���{�^���R�[���o�b�N�Ǘ�(�t�H�[���Ȃ�)
'// �����F       ���{������̃R�[���o�b�N�������ǂ�
'//              �����ꂽ�R���g���[����ID����ɏ������Ăяo���B
'// �����F       control �ΏۃR���g���[��
'// ////////////////////////////////////////////////////////////////////////////
Public Sub ribbonCallback_Copy2CB(control As IRibbonControl)
    Select Case control.ID
        Case "Copy2Clipboard"               '// �Œ蒷�R�s�[
            Call psCopyToClipboard
        Case "Copy2CBMarkdown"               '// �Œ蒷�R�s�[
            Call psCopyToCB_Markdown
    End Select
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// ���\�b�h�F   �N���b�v�{�[�h�փR�s�[
'// �����F       �I��͈͂��Œ蒷�ɐ��`���ăN���b�v�{�[�h�Ɋi�[����B
'// �����F       �Ȃ�
'// �߂�l�F     �Ȃ�
'// ////////////////////////////////////////////////////////////////////////////
Public Sub psCopyToClipboard()
On Error GoTo ErrorHandler
    Const MAX_LEN   As Integer = 80
    Dim tRange      As udTargetRange
    Dim colLen()    As Integer
    Dim idxRow      As Long
    Dim idxCol      As Integer
    Dim bffText     As String
    Dim rslt        As String
    Dim bffHead     As String
    Dim idxArry     As Integer
    Dim textLen     As Integer
  
    '// �I��͈͂��P��ł��邱�Ƃ̊m�F
    If Selection.Areas.Count > 1 Then
        Call MsgBox(MSG_TOO_MANY_RANGE, vbOKOnly, APP_TITLE)
        Exit Sub
    End If
  
    tRange = gfGetTargetRange(ActiveSheet, Selection)
  
    If (tRange.minRow > tRange.maxRow) Or (tRange.minCol > tRange.maxCol) Then
        Call MsgBox(MSG_INVALID_RANGE, vbOKOnly, APP_TITLE)
        Exit Sub
    End If
  
    '// �Z���̒������m�F colLen�Ɋi�[
    idxArry = 0
    For idxCol = tRange.minCol To tRange.maxCol
        ReDim Preserve colLen(idxArry + 1)
        For idxRow = tRange.minRow To tRange.maxRow
            textLen = LenB(StrConv(WorksheetFunction.Clean(Cells(idxRow, idxCol).Text), vbFromUnicode))
            If textLen > colLen(idxArry) Then
              colLen(idxArry) = textLen
            End If
        Next
        colLen(idxArry) = IIf(colLen(idxArry) = 0, 1, colLen(idxArry))
        colLen(idxArry) = IIf(colLen(idxArry) > MAX_COL_LEN, MAX_COL_LEN, colLen(idxArry))  '// 80�o�C�g�ȏ�̒����͐؂�̂�
        idxArry = idxArry + 1
    Next
  
    For idxRow = tRange.minRow To tRange.maxRow
        For idxCol = 0 To tRange.Columns - 1
            bffText = Trim(WorksheetFunction.Clean(Cells(idxRow, idxCol + tRange.minCol).Text)) '// ���s�폜��Trim�B�O���̃g�����͐��l�^�̏ꍇ�̕����p�󔒏����̂��ߕK�v
            bffText = StrConv(LeftB(StrConv(bffText, vbFromUnicode), 80), vbUnicode) '// �ő啶�����ȏ�𑫂���
            textLen = LenB(Trim$(StrConv(bffText, vbFromUnicode)))
            If textLen > MAX_LEN Then    '// 80�����ȏ�͐؂�̂�
                bffText = StrConv(LeftB(StrConv(bffText, vbFromUnicode), colLen(idxCol)), vbUnicode)
            ElseIf IsNumeric(bffText) Or IsDate(bffText) Or pfIsPercentage(bffText) Then    '// ���l�A���t�͉E��
                bffText = Space(colLen(idxCol) - LenB(StrConv(bffText, vbFromUnicode))) & bffText
            Else
                bffText = bffText & Space(colLen(idxCol) - LenB(StrConv(bffText, vbFromUnicode)))
            End If
            rslt = rslt & bffText & Space(1)
        Next
        rslt = Left(rslt, Len(rslt) - 1) & vbCrLf
    Next

    '// �擪�Ɩ����Ɍr����ǉ�
    For idxCol = 0 To tRange.Columns - 1
        bffHead = bffHead & String(colLen(idxCol), "-") & IIf(idxCol = tRange.Columns - 1, vbCrLf, " ")
    Next
    rslt = bffHead & rslt & bffHead
    
    '// �N���b�v�{�[�h�փR�s�[ ��Win10����DataObject�����삵�Ȃ��Ȃ邽�߁A�����SetClip�ɒu������
    Call psSetClip(rslt)
'  Set objData = New DataObject
'  Call objData.SetText(rslt)
'  Call objData.PutInClipboard
'  Set objData = Nothing
  
    Exit Sub
ErrorHandler:
    Call gsShowErrorMsgDlg("mdlCommon.psCopyToClipboard", Err)
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// ���\�b�h�F   �N���b�v�{�[�h��Markdown�`���ŃR�s�[
'// �����F       �I��͈͂��}�[�N�_�E���`���ŃN���b�v�{�[�h�Ɋi�[����B
'// �����F       �Ȃ�
'// �߂�l�F     �Ȃ�
'// ////////////////////////////////////////////////////////////////////////////
Public Sub psCopyToCB_Markdown()
On Error GoTo ErrorHandler
    Dim tRange      As udTargetRange
    Dim colLen()    As Integer
    Dim idxRow      As Long
    Dim idxCol      As Integer
    Dim bffText     As String
    Dim rslt        As String
    Dim bffHead     As String
    Dim idxArry     As Integer
    Dim textLen     As Integer
  
    '// �I��͈͂��P��ł��邱�Ƃ̊m�F
    If Selection.Areas.Count > 1 Then
        Call MsgBox(MSG_TOO_MANY_RANGE, vbOKOnly, APP_TITLE)
        Exit Sub
    End If
  
    tRange = gfGetTargetRange(ActiveSheet, Selection)
  
    If (tRange.minRow > tRange.maxRow) Or (tRange.minCol > tRange.maxCol) Then
        Call MsgBox(MSG_INVALID_RANGE, vbOKOnly, APP_TITLE)
        Exit Sub
    End If
    
    '// �w�b�_�̏o��
    rslt = "|"
    For idxCol = 0 To tRange.Columns - 1
        rslt = rslt & " " & Replace(Cells(tRange.minRow, idxCol + tRange.minCol).Text, vbLf, "<br>") & " |"
    Next
    rslt = rslt & vbCrLf & "|"
    For idxCol = 0 To tRange.Columns - 1
        Select Case Cells(tRange.minRow, idxCol + tRange.minCol).HorizontalAlignment
            Case xlRight
                rslt = rslt & " " & "-: |"
            Case xlCenter
                rslt = rslt & " " & ":-: |"
            Case Else
                rslt = rslt & " " & "- |"
        End Select
    Next
    
    '// �f�[�^�s�̏o��
    For idxRow = tRange.minRow + 1 To tRange.maxRow
        rslt = rslt & vbCrLf & "|"
        For idxCol = 0 To tRange.Columns - 1
            rslt = rslt & " " & Replace(Cells(idxRow, idxCol + tRange.minCol).Text, vbLf, "<br>") & " |"
        Next
    Next
    
    '// �N���b�v�{�[�h�փR�s�[ ��Win10����DataObject�����삵�Ȃ��Ȃ邽�߁A�����SetClip�ɒu������
    Call psSetClip(rslt)
    
    Exit Sub
ErrorHandler:
    Call gsShowErrorMsgDlg("mdlCommon.psCopyToClipboard_MarkDown", Err)
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// �����F       Win10 ���� DataObject.PutInClipboard �������Ȃ��Ȃ������߁A�����Ƃ��ăe�L�X�g�{�b�N�X���o�R���ăR�s�[
'// �����F       �R�s�[�Ώە�����
'// ////////////////////////////////////////////////////////////////////////////
Private Sub psSetClip(bffText As String)
    With CreateObject("Forms.TextBox.1")
        .MultiLine = True
        .Text = bffText
        .SelStart = 0
        .SelLength = .TextLength
        .Copy
    End With
    DoEvents    '// ��Q����̂��߁A��xOS�ɏ�����߂��i�Č����Ⴂ���߂��̑Ώ����ǂ����͊m�ؖ����j
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// �����F       �����̕����񂪃p�[�Z���g�`�����𔻒肷��
'// �����F       �R�s�[�Ώە�����
Private Function pfIsPercentage(bffText As String) As Boolean
    If bffText = BLANK Then
        pfIsPercentage = False
    ElseIf Right(bffText, 1) = "%" And IsNumeric(Left(bffText, Len(bffText) - 1)) Then
        pfIsPercentage = True
    Else
        pfIsPercentage = False
    End If
End Function


'// ////////////////////////////////////////////////////////////////////////////
'// END
'// ////////////////////////////////////////////////////////////////////////////