Attribute VB_Name = "mdlGroup"
'// ////////////////////////////////////////////////////////////////////////////
'// �v���W�F�N�g   : �g���c�[��
'// �^�C�g��       : �I�u�W�F�N�g�̕␳�@�\
'// ���W���[��     : mdlAdjustShape
'// ����           : ���R�l�N�^��u���b�N���Ȃǂ̃I�u�W�F�N�g�̔������@�\
'//                  ����mdlFeatures�iV2.1.1�܂Łj
'// ////////////////////////////////////////////////////////////////////////////
'// Copyright (c) by Koichiro.
'// ////////////////////////////////////////////////////////////////////////////
Option Explicit
Option Base 0


'// ////////////////////////////////////////////////////////////////////////////
'// ���\�b�h�F   ���{���{�^���R�[���o�b�N�Ǘ�(�t�H�[���Ȃ�)
'// �����F       ���{������̃R�[���o�b�N�������ǂ�
'//              �����ꂽ�R���g���[����ID����ɏ������Ăяo���B
'// �����F       control �ΏۃR���g���[��
'// ////////////////////////////////////////////////////////////////////////////
Public Sub ribbonCallback_Group(control As IRibbonControl)
    Select Case control.ID
        Case "groupRow"                     '// �O���[�v�� �s
            Call psSetGroup_Row
        Case "groupCol"                     '// �O���[�v�� ��
            Call psSetGroup_Col
        Case "removeDup"                    '// �d���̃J�E���g
            Call psDistinctVals
        Case "listDup"                      '// �d�����K�w���ɕ␳
            Call psGroupVals
    End Select
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// ���\�b�h�F   �O���[�v�ݒ�i�s�j
'// �����F       �O���[�v�������ݒ肷��B
'// �����F       �Ȃ�
'// �߂�l�F     �Ȃ�
'// ////////////////////////////////////////////////////////////////////////////
Private Sub psSetGroup_Row()
    Dim idxStart    As Long
    Dim idxEnd      As Long
    Dim idxRow      As Long
    Dim idxCol      As Integer
    Dim tRange      As udTargetRange
    Dim childRange  As Range
    
    '// ���O�`�F�b�N�i�A�N�e�B�u�V�[�g�ی�A�I���^�C�v���Z���j
    If Not gfPreCheck(protectCont:=True, selType:=TYPE_RANGE) Then
        Exit Sub
    End If
    
'    Application.ScreenUpdating = False
    Call gsSuppressAppEvents
    
  
    '// �A�E�g���C���̏W�v�ʒu��ύX
    With ActiveSheet.Outline
        .SummaryRow = xlAbove
    End With
  
        '// �O���[�v�ݒ�
        For Each childRange In Selection.Areas
            tRange = gfGetTargetRange(ActiveSheet, childRange)
            
            idxStart = 0
            idxEnd = 0
            idxCol = tRange.minCol
            
            For idxRow = tRange.minRow To tRange.maxRow
                If idxStart = 0 Then
                    idxStart = idxRow + 1
                    idxEnd = idxRow + 1
                ElseIf Trim(Cells(idxRow, idxCol).Text) = BLANK Then
                    idxEnd = idxRow
                ElseIf Trim(Cells(idxRow - 1, idxCol).Text) = BLANK Then
                    Range(Cells(idxStart, 1), Cells(idxEnd, 1)).Rows.Group
                    idxStart = idxRow + 1
                    idxEnd = idxRow + 1
                Else
                    idxStart = idxRow + 1
                    idxEnd = idxRow + 1
                End If
            Next
            If idxStart < idxEnd Then
                Range(Cells(idxStart, 1), Cells(idxEnd, 1)).Rows.Group
            End If
      Next
      
      Call gsResumeAppEvents
'      Application.ScreenUpdating = True
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// ���\�b�h�F   �O���[�v�ݒ�i��j
'// �����F       �O���[�v�������ݒ肷��B
'// �����F       �Ȃ�
'// �߂�l�F     �Ȃ�
'// ////////////////////////////////////////////////////////////////////////////
Private Sub psSetGroup_Col()
    Dim idxStart    As Long
    Dim idxEnd      As Long
    Dim idxRow      As Long
    Dim idxCol      As Integer
    Dim tRange      As udTargetRange
    Dim childRange  As Range
    
    '// ���O�`�F�b�N�i�A�N�e�B�u�V�[�g�ی�A�I���^�C�v���Z���j
    If Not gfPreCheck(protectCont:=True, selType:=TYPE_RANGE) Then
        Exit Sub
    End If
    
    Call gsSuppressAppEvents
'    Application.ScreenUpdating = False
    
    '// �A�E�g���C���̏W�v�ʒu��ύX
    With ActiveSheet.Outline
        .SummaryColumn = xlLeft
    End With
    
    '// �O���[�v�ݒ�
    For Each childRange In Selection.Areas
        tRange = gfGetTargetRange(ActiveSheet, childRange)
        
        idxStart = 0
        idxEnd = 0
        idxRow = tRange.minRow
        
        For idxCol = tRange.minCol To tRange.maxCol
            If idxStart = 0 Then
                idxStart = idxCol + 1
                idxEnd = idxCol + 1
            ElseIf Trim(Cells(idxRow, idxCol).Text) = BLANK Then
                idxEnd = idxCol
            ElseIf Trim(Cells(idxRow, idxCol - 1).Text) = BLANK Then
                Range(Cells(1, idxStart), Cells(1, idxEnd)).Columns.Group
                idxStart = idxCol + 1
                idxEnd = idxCol + 1
            Else
                idxStart = idxCol + 1
                idxEnd = idxCol + 1
            End If
        Next
        If idxStart < idxEnd Then
            Range(Cells(1, idxStart), Cells(1, idxEnd)).Columns.Group
        End If
    Next
    
'    Application.ScreenUpdating = True
    Call gsResumeAppEvents
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// ���\�b�h�F   �l�̏d����r�����Ĉꗗ�i�J�E���g�j
'// �����F       �d���l��r������B
'// ////////////////////////////////////////////////////////////////////////////
Private Sub psDistinctVals()
    Dim idxRow      As Long
    Dim idxCol      As Integer
    Dim tRange      As udTargetRange
    
    Dim bff         As Variant
    Dim dict        As Object
    Dim keyString   As String
    Dim keyArray()  As String
    Dim resultSheet As Worksheet
    
    '// �Z�����I������Ă��邱�Ƃ��`�F�b�N
    If TypeName(Selection) <> TYPE_RANGE Then
        Call MsgBox(MSG_NOT_RANGE_SELECT, vbOKOnly, APP_TITLE)
        Exit Sub
    End If
    
    '// �`�F�b�N
    If Selection.Areas.Count > 1 Then
        Call MsgBox(MSG_TOO_MANY_RANGE, vbOKOnly, APP_TITLE)
        Exit Sub
    End If
    
    Call gsSuppressAppEvents
    
    tRange = gfGetTargetRange(ActiveSheet, Selection)
    
    bff = Selection.Areas(1)
    Set dict = CreateObject("Scripting.Dictionary")
    
    For idxRow = 1 To tRange.maxRow - tRange.minRow + 1
        '// �s�̃Z�����������ĕ���������
        keyString = BLANK
        For idxCol = 1 To tRange.maxCol - tRange.minCol + 1
            If Not IsError(bff(idxRow, idxCol)) Then
                keyString = keyString & Chr(127) & bff(idxRow, idxCol)
            End If
        Next
        
        If Not dict.Exists(keyString) Then
            Call dict.Add(keyString, "1")
        Else
            dict.Item(keyString) = CStr(CLng(dict.Item(keyString)) + 1)
        End If
    Next
    
    '// ���ʏo��
    Call Workbooks.Add
    Set resultSheet = ActiveWorkbook.ActiveSheet
    
    '// �w�b�_�̐ݒ�B�u�J�E���g�v�̃w�b�_�ʒu�����킹�邽�߁AHDR_DISTINCT����"@"��񐔂ɍ��킹��Replace����
    Call gsDrawResultHeader(resultSheet, Replace(HDR_DISTINCT, "@", String(tRange.Columns, ";")), 1)
    
    '// �L�[�̔z���variant�Ɋi�[
    bff = dict.Keys
    
    For idxRow = 0 To dict.Count - 1
        keyArray = Split(bff(idxRow), Chr(127))  '// split�͓Y�����P����J�n�̎d�l�H
        For idxCol = 1 To UBound(keyArray)
            resultSheet.Cells(idxRow + 2, idxCol).Value = keyArray(idxCol)
        Next
        
        resultSheet.Cells(idxRow + 2, tRange.maxCol - tRange.minCol + 2).Value = dict.Item(bff(idxRow))
    Next
    
    '//�t�H���g
    Call resultSheet.Cells.Select
    Selection.Font.Name = APP_FONT
    Selection.Font.Size = APP_FONT_SIZE
    
    '// �r���`��
    Call gsPageSetup_Lines(resultSheet, 1)
    
    '// ����Ƃ��ɕۑ������߂Ȃ�
    ActiveWorkbook.Saved = True
    
    Set dict = Nothing
    
    Call gsResumeAppEvents
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// ���\�b�h�F   �l���K�w���ɕ␳����
'// �����F       �d���l���K�w���ɕ␳����
'// ////////////////////////////////////////////////////////////////////////////
Private Sub psGroupVals()
    Dim idxRow        As Long
    Dim idxCol        As Integer
    Dim tRange        As udTargetRange
    Dim aryIdx        As Integer
    Dim aryLastVal(8) As String
    
    '// ���O�`�F�b�N�i�A�N�e�B�u�V�[�g�ی�A�I���^�C�v���Z���j
    If Not gfPreCheck(protectCont:=True, selType:=TYPE_RANGE) Then
        Exit Sub
    End If
    
'    '// �Z�����I������Ă��邱�Ƃ��`�F�b�N
'    If TypeName(Selection) <> TYPE_RANGE Then
'        Call MsgBox(MSG_NOT_RANGE_SELECT, vbOKOnly, APP_TITLE)
'        Exit Sub
'    End If
    
    '// �`�F�b�N
    If Selection.Areas.Count > 1 Then
        Call MsgBox(MSG_TOO_MANY_RANGE, vbOKOnly, APP_TITLE)
        Exit Sub
    ElseIf Selection.Columns.Count > 8 Then
        Call MsgBox(MSG_TOO_MANY_COLS_8, vbOKOnly, APP_TITLE)
        Exit Sub
    End If
    
    Call gsSuppressAppEvents
    
    tRange = gfGetTargetRange(ActiveSheet, Selection)
    
    For idxRow = tRange.minRow To tRange.maxRow
        For idxCol = tRange.minCol To tRange.maxCol
            If (aryLastVal(idxCol - tRange.minCol) = BLANK) Or (aryLastVal(idxCol - tRange.minCol) <> Cells(idxRow, idxCol).Text) Then
                '// ���O�̒l���قȂ�ꍇ (�� ���O�̒l���u�����N�̏ꍇ)
                '// �z���̃��x���̒��O�̒l���N���A
                For aryIdx = tRange.Columns To idxCol Step -1
                    aryLastVal(aryIdx - 1) = BLANK
                Next
                aryLastVal(idxCol - tRange.minCol) = Cells(idxRow, idxCol).Text
            Else
                Cells(idxRow, idxCol).Value = BLANK
            End If
        Next
    Next
    
    Call gsResumeAppEvents
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// END
'// ////////////////////////////////////////////////////////////////////////////

