Attribute VB_Name = "mdlHyperlink"
'// ////////////////////////////////////////////////////////////////////////////
'// �v���W�F�N�g   : �g���c�[��
'// �^�C�g��       : �n�C�p�[�����N
'// ���W���[��     : mdlHyperlink
'// ����           : �n�C�p�[�����N�̐ݒ肨��щ����@�\
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
Public Sub ribbonCallback_Hyperlink(control As IRibbonControl)
    Select Case control.ID
        Case "AddLink"                      '// �����N�̒ǉ�
            Call psSetHyperLink
        Case "RemoveLink"                   '// �����N�̍폜
            Call psRemoveHyperLink
    End Select
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// ���\�b�h�F   �n�C�p�[�����N�̐ݒ�
'// �����F       �I��͈͂̃n�C�p�[�����N��ݒ肷��
'//              �W���@�\�̃n�C�p�[�����N�ݒ�ł̓e�L�X�g�������ς�邽�߁A�ݒ�O�̏�����ێ�����
'// �����F       �Ȃ�
'// �߂�l�F     �Ȃ�
'// ////////////////////////////////////////////////////////////////////////////
Private Sub psSetHyperLink()
    Dim tCell     As Range    '// �ΏۃZ��
'    Dim tRange    As udTargetRange
'    Dim childRange As Range
'    Dim idxRow    As Long
'    Dim idxCol    As Integer
    Dim fontName  As String
    Dim fontSize  As String
    Dim fontBold  As Boolean
    Dim fontItlic As Boolean
    Dim fontColor As Double
    
    '// ���O�`�F�b�N�i�A�N�e�B�u�V�[�g�ی�A�I���^�C�v���Z���j
    If Not gfPreCheck(protectCont:=True, selType:=TYPE_RANGE) Then
        Exit Sub
    End If
    
    Call gsSuppressAppEvents
    
    For Each tCell In Selection.SpecialCells(xlCellTypeConstants, xlNumbers + xlTextValues)
        If Trim(tCell.Text) <> BLANK Then
            fontName = tCell.Font.Name
            fontSize = tCell.Font.Size
            fontBold = tCell.Font.Bold
            fontItlic = tCell.Font.Italic
            fontColor = tCell.Font.Color
            Call tCell.Hyperlinks.Add(Anchor:=tCell, Address:=tCell.Text)
            tCell.Font.Name = fontName
            tCell.Font.Size = fontSize
            tCell.Font.Bold = fontBold
            tCell.Font.Italic = fontItlic
            tCell.Font.Color = fontColor
        End If
    Next
    
'    For Each childRange In Selection.Areas
'        tRange = gfGetTargetRange(ActiveSheet, childRange)
'        For idxRow = tRange.minRow To tRange.maxRow
'            For idxCol = tRange.minCol To tRange.maxCol
'                If Trim(Cells(idxRow, idxCol).Text) <> BLANK Then
'                    fontName = Cells(idxRow, idxCol).Font.Name
'                    fontSize = Cells(idxRow, idxCol).Font.Size
'                    fontBold = Cells(idxRow, idxCol).Font.Bold
'                    fontItlic = Cells(idxRow, idxCol).Font.Italic
'                    fontColor = Cells(idxRow, idxCol).Font.Color
'                    Call Cells(idxRow, idxCol).Hyperlinks.Add(Anchor:=Cells(idxRow, idxCol), Address:=Cells(idxRow, idxCol).Text)
'                    Cells(idxRow, idxCol).Font.Name = fontName
'                    Cells(idxRow, idxCol).Font.Size = fontSize
'                    Cells(idxRow, idxCol).Font.Bold = fontBold
'                    Cells(idxRow, idxCol).Font.Italic = fontItlic
'                    Cells(idxRow, idxCol).Font.Color = fontColor
'                End If
'            Next
'        Next
'    Next
    Call gsResumeAppEvents
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// ���\�b�h�F   �n�C�p�[�����N�̍폜
'// �����F       �I��͈͂̃n�C�p�[�����N���폜����
'// �����F       �Ȃ�
'// �߂�l�F     �Ȃ�
'// Excel 2010���_�ŁuHyperLink�̃N���A�v���W����������Ă��邪�A�c�[���Ƃ���UI���c�����ƂƂ����B
'// ////////////////////////////////////////////////////////////////////////////
Private Sub psRemoveHyperLink()
'    Dim tRange    As udTargetRange
'    Dim idxRow    As Long
'    Dim idxCol    As Integer
'    Dim fontName  As String
'    Dim fontSize  As String
'    Dim borderLines(8, 3) As Long
'    Dim childRange As Range
    
    '// ���O�`�F�b�N�i�A�N�e�B�u�V�[�g�ی�A�I���^�C�v���Z���j
    If Not gfPreCheck(protectCont:=True, selType:=TYPE_RANGE) Then
        Exit Sub
    End If
    
    Call Selection.ClearHyperlinks
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// END
'// ////////////////////////////////////////////////////////////////////////////

