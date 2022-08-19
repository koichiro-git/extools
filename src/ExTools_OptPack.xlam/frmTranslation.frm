VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTranslation 
   Caption         =   "�|��"
   ClientHeight    =   1860
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4605
   OleObjectBlob   =   "frmTranslation.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "frmTranslation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'// ////////////////////////////////////////////////////////////////////////////
'// �v���W�F�N�g   : �g���c�[�� �ǉ��p�b�N
'// �^�C�g��       : �|��t�H�[��
'// ���W���[��     : frmTranslation
'// ����           : �I��͈͂̃Z���������|�󂷂�
'// ////////////////////////////////////////////////////////////////////////////
'// Copyright (c) by Koichiro.
'// ////////////////////////////////////////////////////////////////////////////
Option Explicit
Option Base 0

'// ////////////////////////////////////////////////////////////////////////////
'// �v���C�x�[�g�萔
Private Const URL_DEEPL_FREE = "https://api-free.deepl.com/v2/translate?"
Private Const URL_DEEPL_LICENSED = "https://api.deepl.com/v2/translate?"

Private Const SEPARATOR = "--" + vbLf
Private Const TAG_LINE_FEED = "<LF/>"

'// �v���C�x�[�g�ϐ�
Private auth_key        As String
Private license         As String
Private url_deepl       As String

'// //////////////////////////////////////////////////////////////////
'// �C�x���g�F �t�H�[�� ��������
Private Sub UserForm_Initialize()
    '// �ݒ�t�@�C������L�[��ǂݍ���
    auth_key = gfGetIniFileSetting("TRANSLATE", "DEEPL_AUTH_KEY")   '// �F�؃L�[
    license = gfGetIniFileSetting("TRANSLATE", "DEEPL_LICENSE")     '// ���C�Z���X�iFREE/PRO�j
    If license <> "FREE" Then
        url_deepl = URL_DEEPL_LICENSED
    Else
        url_deepl = URL_DEEPL_FREE
    End If
    
    '// �R���{�{�b�N�X�ݒ�
    Call gsSetCombo(cmbLanguage, CMB_TRN_LANGUAGE, 0)
    Call gsSetCombo(cmbOutput, CMB_TRN_OUTPUT, 0)

    '// �L���v�V�����ݒ�
    frmTranslation.Caption = IIf(license <> "FREE", BLANK, LBL_TRN_FORM & " " & pfGetServiceUsage)
    lblLanguage.Caption = LBL_TRN_LANG
    lblOutput.Caption = LBL_TRN_OUTPUT
End Sub


'// //////////////////////////////////////////////////////////////////
'// �C�x���g�F ����{�^�� �N���b�N��
Private Sub cmdClose_Click()
    Call Me.Hide
End Sub


'// //////////////////////////////////////////////////////////////////
'// �C�x���g�F ���s�{�^�� �N���b�N��
Private Sub cmdExecute_Click()
    Dim tCells      As Range    '// �ϊ��ΏۃZ��
    
    '// ���O�`�F�b�N�i�A�N�e�B�u�V�[�g�ی�A�I���^�C�v���Z���j
    If Not gfPreCheck(protectCont:=True, selType:=TYPE_RANGE) Then
        Exit Sub
    End If
    Call gsSuppressAppEvents
    
    If Selection.Count > 1 Then
        Set tCells = Selection.SpecialCells(xlCellTypeConstants, xlNumbers + xlTextValues)
    Else
        Set tCells = ActiveCell
    End If
    
    Call psTranslate_DeepL(tCells)  '// DeepL
    
    Call gsResumeAppEvents
    Call Me.Hide
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// ���\�b�h�F   �|�� ��֐� (DeepL)
'// �����F       �I��͈͂̕������|�󂷂�
'// �����F       tCells �Ώ۔͈�
'// ////////////////////////////////////////////////////////////////////////////
Private Sub psTranslate_DeepL(tCells As Range)
'On Error GoTo ErrorHandler
    Dim httpReq         As New XMLHTTP60
    Dim htmlDoc         As New HTMLDocument
    Dim reqParam        As String
    Dim sourceText      As String
    Dim resultText      As String
    Dim appendIdx       As Integer
    
    Dim tCell           As Range
    
    For Each tCell In tCells
        sourceText = EncodeURL2(tCell.Text, htmlDoc)
        reqParam = "&target_lang=" & cmbLanguage.Value & _
                   "&auth_key=" & auth_key & _
                   "&text=" & sourceText & _
                   "&tag_handling=xml"
        
        Call httpReq.Open("POST", url_deepl)
        Call httpReq.setRequestHeader("Content-Type", "application/x-www-form-urlencoded")
        Call httpReq.send(reqParam)
        
        Application.StatusBar = tCell.Address & ": " & Left(tCell.Text, 30)
        Do While httpReq.readyState < 4
            DoEvents
        Loop
        
        '// ���X�|���X�擾��iJSON���C�u�����͎g�p�����A������Ƃ��ď�������j
        If httpReq.Status = 200 Then
            '// JSON��茋�ʕ����񒊏o�B�`���� {"translations":[{"detected_source_language":"EN","text":"������"}]}
            resultText = Mid(httpReq.responseText, InStr(1, httpReq.responseText, """text"":") + 8)     '// "text":" ������擾
            resultText = Left(resultText, Len(resultText) - 4)  '// �Ō�� "}]} ���폜
            resultText = DecodeResultText(resultText)
        Else
            resultText = "Error: HTTP Response=" & httpReq.Status
        End If
        
        '// ���ʕ������ݒ�
        Select Case cmbOutput.Value
            Case 0  '// �����̉�
                appendIdx = Len(tCell.Text) + 2 + Len(SEPARATOR)    '// ���s�R�[�h�ƃZ�p���[�^(--)�̕����J�nIndex�Ƃ���SEPARATOR�͉��s�R�[�h�܁B
                tCell.Value = tCell.Text & vbLf & SEPARATOR & resultText
                tCell.Characters(Start:=appendIdx, Length:=Len(resultText)).Font.Color = RGB(0, 0, 255)
            Case 1  '// �����̏�
                appendIdx = Len(tCell.Text)
                tCell.Value = resultText & vbLf & SEPARATOR & vbLf & tCell.Text
                tCell.Characters(Start:=1, Length:=Len(resultText)).Font.Color = RGB(0, 0, 255)
            Case 2  '// �������㏑��
                tCell.Value = resultText
            Case 3  '// �E�̃Z���ɏ㏑��
                tCell.Offset(, 1).Value = resultText
            Case 4  '// �R�����g�ɐݒ�
                Call tCell.NoteText(resultText)
                tCell.Comment.Shape.TextFrame.AutoSize = True
                '// �R�����g������
                If tCell.Comment.Shape.Width > 500 Then
                    tCell.Comment.Shape.Height = tCell.Comment.Shape.Height * (1 / (500 / tCell.Comment.Shape.Width)) * 1.2
                    tCell.Comment.Shape.Width = 500
                End If
        End Select
    Next
    
'    httpReq.Close
    Set httpReq = Nothing
    
    Exit Sub
ErrorHandler:
    Call gsResumeAppEvents
    Call gsShowErrorMsgDlg_VBA("frmTranslate.psTranslate_DeepL", Err)
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// ���\�b�h�F   URL�G���R�[�h
'// �����F       �����̕�������G���R�[�h����B
'//              Excel2013�ȑO�ł����s�\�ɂ��邽��WorksheetFunction.EncodeURL�͎g�p���Ȃ�
'// ////////////////////////////////////////////////////////////////////////////
Private Function EncodeURL2(targetTxt As String, htmlDoc As HTMLDocument) As String
    Dim elm         As HTMLHtmlElement
    
    targetTxt = Replace(targetTxt, "\", "\\")
'    targetTxt = Replace(targetTxt, "#", "\#")
    targetTxt = Replace(targetTxt, "'", "\'")
    targetTxt = Replace(targetTxt, vbLf, TAG_LINE_FEED)
    'Set htmlDoc = CreateObject("htmlfile")
    
    Set elm = htmlDoc.createElement("span")
    Call elm.setAttribute("id", "result")
    Call htmlDoc.body.appendChild(elm)
    Call htmlDoc.parentWindow.execScript("document.getElementById('result').innerText = encodeURIComponent('" & targetTxt & "');", "JScript")
    EncodeURL2 = elm.innerText
    Call elm.ParentNode.RemoveChild(elm)
    Set elm = Nothing
End Function


'// ////////////////////////////////////////////////////////////////////////////
'// ���\�b�h�F   ���ʕ�����␳
'// �����F       �|��T�[�r�X����߂��ꂽ������̃^�O�Ȃǂ�␳����
'// ////////////////////////////////////////////////////////////////////////////
Private Function DecodeResultText(targetTxt As String) As String
    DecodeResultText = Replace(targetTxt, TAG_LINE_FEED, vbLf)
End Function


'// ////////////////////////////////////////////////////////////////////////////
'// ���\�b�h�F   �T�[�r�X�g�p�󋵎擾
'// �����F       �|��T�[�r�X���烂�j�^�����O�󋵂��擾���A������Ŗ߂�
'// ////////////////////////////////////////////////////////////////////////////
Private Function pfGetServiceUsage() As String
On Error GoTo ErrorHandler
    Dim httpReq         As New XMLHTTP60
    Dim reqParam        As String
    Dim sourceText      As String
    Dim resultText      As String
    
    Call httpReq.Open("POST", "https://api-free.deepl.com/v2/usage?")
    Call httpReq.setRequestHeader("Content-Type", "application/x-www-form-urlencoded")
    reqParam = "&auth_key=" & auth_key
    Call httpReq.send(reqParam)
    
    Do While httpReq.readyState < 4
        DoEvents
    Loop
        
    '// ���X�|���X�擾��iVBA-JSON�͎g�p�����A������Ƃ��ď�������j
    If httpReq.Status = 200 Then
        pfGetServiceUsage = pfFormatUsage(httpReq.responseText)
    Else
        pfGetServiceUsage = BLANK
    End If
    
'    httpReq.Close
    Set httpReq = Nothing
    Exit Function
    
ErrorHandler:
    Call gsShowErrorMsgDlg_VBA("frmTranslate.pfGetServiceUsage", Err)
End Function


'// ////////////////////////////////////////////////////////////////////////////
'// ���\�b�h�F   DeepL�g�p��JSON�t�H�[�}�b�g
'// �����F       ������JSON��������A 999 / 999 �̌`���̕�����Ŗ߂��B
'//              pfGetServiceUsage�̃T�u�v���V�[�W��
'//              ���K�\���Ő��l�ȊO�����ׂĎ�菜����r���̃J���}���X���b�V���ɒu��������
'// �����F       str: JSON������
'// ////////////////////////////////////////////////////////////////////////////
Private Function pfFormatUsage(str As String) As String
    Dim reg     As Object
    
    Set reg = CreateObject("VBScript.RegExp")
    reg.Pattern = "[^0-9,]"
    reg.Global = True
    pfFormatUsage = reg.Replace(StrConv(str, vbNarrow), "")
    pfFormatUsage = Replace(pfFormatUsage, ",", " / ")
End Function


'// ////////////////////////////////////////////////////////////////////////////
'// END
'// ////////////////////////////////////////////////////////////////////////////

