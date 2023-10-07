Attribute VB_Name = "mdlDrawLines"
'// ////////////////////////////////////////////////////////////////////////////
'// �v���W�F�N�g   : �g���c�[��
'// �^�C�g��       : �r���`��@�\
'// ���W���[��     : mdlDrawLines
'// ����           : �s�E��̃w�b�_����уf�[�^�����̌r����`��
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
Public Sub ribbonCallback_DrawLines(control As IRibbonControl)
    Select Case control.ID
        Case "BorderRowHead"                '// �s�w�b�_�̌r��
            Call gsDrawLine_Header
        Case "BorderColHead"                '// ��w�b�_�̌r��
            Call gsDrawLine_Header_Vert
        Case "BorderData"                   '// �f�[�^�̈�̌r��
            Call gsDrawLine_Data
    End Select
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// ���\�b�h�F   �r���`��i�w�b�_�j
'// �����F       �w�b�_���̌r����`�悷��i���j
'// ////////////////////////////////////////////////////////////////////////////
Public Sub gsDrawLine_Header()
    Dim baseRow As Long     '// �I��̈�̊J�n�ʒu
    Dim baseCol As Integer  '// �I��̈�̊J�n�ʒu
    Dim selRows As Long     '// �I��̈�̍s��
    Dim selCols As Integer  '// �I��̈�̗�
    Dim idxRow  As Long
    Dim idxCol  As Integer
    Dim offRow  As Long
    Dim offCol  As Integer
    Dim childRange As Range
    
    '// ���O�`�F�b�N�i�A�N�e�B�u�V�[�g�ی�A�I���^�C�v���Z���j
    If Not gfPreCheck(protectCont:=True, selType:=TYPE_RANGE) Then
        Exit Sub
    End If
    
    Call gsSuppressAppEvents
    
    For Each childRange In Selection.Areas
        '// �r�����N���A
        childRange.Borders.LineStyle = xlNone
        childRange.Borders(xlDiagonalDown).LineStyle = xlNone
        childRange.Borders(xlDiagonalUp).LineStyle = xlNone
        
        '// �I��͈͂̊J�n�E�I���ʒu�擾
        baseRow = childRange.Row
        baseCol = childRange.Column
        selRows = childRange.Rows.Count
        selCols = childRange.Columns.Count
        
        For idxRow = baseRow To baseRow + selRows
            For idxCol = baseCol To baseCol + selCols
                offRow = 0
                offCol = 0
                If (Cells(idxRow, idxCol).Text <> BLANK) Or ((idxRow = baseRow) And (idxCol = baseCol)) Then
                    For offRow = idxRow To baseRow + selRows - 1
                        If (offRow = idxRow) Or Cells(offRow, idxCol).Value = BLANK Then
                            Cells(offRow, idxCol).Borders(xlEdgeLeft).LineStyle = xlContinuous
                        Else
                            Exit For
                        End If
                    Next
                    For offCol = idxCol To baseCol + selCols - 1
                        If (offCol = idxCol) Or Cells(idxRow, offCol).Text = BLANK Then
                            Cells(idxRow, offCol).Borders(xlEdgeTop).LineStyle = xlContinuous
                            If Cells(idxRow, offCol).Borders(xlEdgeRight).LineStyle = xlContinuous Then
                                Exit For
                            End If
                        Else
                            Exit For
                        End If
                    Next
                End If
                '// �ő��ɒB�����ꍇ�͏I��
                If idxCol = Columns.Count Then
                    Exit For
                End If
            Next
            '// �ő�s�ɒB�����ꍇ�͏I��
            If idxRow = Rows.Count Then
                Exit For
            End If
        Next
        
        With childRange.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With
        With childRange.Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With
    Next
    
    Call gsResumeAppEvents
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// ���\�b�h�F   �r���`��i�w�b�_�j�F�c
'// �����F       �w�b�_���̌r����`�悷��i�c�j
'// �����F       �Ȃ�
'// �߂�l�F     �Ȃ�
'// ////////////////////////////////////////////////////////////////////////////
Public Sub gsDrawLine_Header_Vert()
    Dim baseRow As Long     '// �I��̈�̊J�n�ʒu
    Dim baseCol As Integer  '// �I��̈�̊J�n�ʒu
    Dim selRows As Long     '// �I��̈�̍s��
    Dim selCols As Integer  '// �I��̈�̗�
    Dim idxRow  As Long
    Dim idxCol  As Integer
    Dim offRow  As Long
    Dim offCol  As Integer
    Dim childRange As Range
  
    '// ���O�`�F�b�N�i�A�N�e�B�u�V�[�g�ی�A�I���^�C�v���Z���j
    If Not gfPreCheck(protectCont:=True, selType:=TYPE_RANGE) Then
        Exit Sub
    End If
    
    Call gsSuppressAppEvents
    
    For Each childRange In Selection.Areas
        '// �r�����N���A
        childRange.Borders.LineStyle = xlNone
        childRange.Borders(xlDiagonalDown).LineStyle = xlNone
        childRange.Borders(xlDiagonalUp).LineStyle = xlNone
        
        '// �I��͈͂̊J�n�E�I���ʒu�擾
        baseRow = childRange.Row
        baseCol = childRange.Column
        selRows = childRange.Rows.Count
        selCols = childRange.Columns.Count
      
        For idxCol = baseCol To baseCol + selCols
            For idxRow = baseRow To baseRow + selRows
                offRow = 0
                offCol = 0
                If (Cells(idxRow, idxCol).Value <> BLANK) Or ((idxRow = baseRow) And (idxCol = baseCol)) Then
                    For offCol = idxCol To baseCol + selCols - 1
                        If (offCol = idxCol) Or Cells(idxRow, offCol).Value = BLANK Then
                            Cells(idxRow, offCol).Borders(xlEdgeTop).LineStyle = xlContinuous
                        Else
                            Exit For
                        End If
                    Next
                    For offRow = idxRow To baseRow + selRows - 1
                        If (offRow = idxRow) Or Cells(offRow, idxCol).Value = BLANK Then
                            Cells(offRow, idxCol).Borders(xlEdgeLeft).LineStyle = xlContinuous
                            If Cells(offRow, idxCol).Borders(xlEdgeBottom).LineStyle = xlContinuous Then
                                Exit For
                            End If
                        Else
                            Exit For
                        End If
                    Next
                End If
                '// �ő�s�ɒB�����ꍇ�͏I��
                If idxRow = Rows.Count Then
                    Exit For
                End If
            Next
            '// �ő��ɒB�����ꍇ�͏I��
            If idxCol = Columns.Count Then
                Exit For
            End If
        Next
    
        With childRange.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With
        With childRange.Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With
    Next
    
    Call gsResumeAppEvents
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// ���\�b�h�F   �r���`��i�f�[�^�j
'// �����F       �f�[�^���̌r����`�悷��
'//              �I��͈͎��ӕ���xlThin�A������xlHairline�ŕ`�悷��
'// �����F       �Ȃ�
'// �߂�l�F     �Ȃ�
'// ////////////////////////////////////////////////////////////////////////////
Public Sub gsDrawLine_Data()
    '// ���O�`�F�b�N�i�A�N�e�B�u�V�[�g�ی�A�I���^�C�v���Z���j
    If Not gfPreCheck(protectCont:=True, selType:=TYPE_RANGE) Then
        Exit Sub
    End If
    
    Call gsSuppressAppEvents
    
    '// �r���`��
    Selection.Borders.LineStyle = xlContinuous
    Selection.Borders.Weight = xlThin
    
    If Selection.Columns.Count > 1 Then
        With Selection.Borders(xlInsideVertical)
            .LineStyle = xlContinuous
            .Weight = xlHairline
        End With
    End If
    
    If Selection.Rows.Count > 1 Then
        With Selection.Borders(xlInsideHorizontal)
            .LineStyle = xlContinuous
            .Weight = xlHairline
        End With
    End If
    
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone

    Call gsResumeAppEvents
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// END
'// ////////////////////////////////////////////////////////////////////////////

