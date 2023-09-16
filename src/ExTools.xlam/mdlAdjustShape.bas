Attribute VB_Name = "mdlAdjustShape"
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
Public Sub ribbonCallback_AdjustShape(control As IRibbonControl)
    Select Case control.ID
        Case "AdjShapeElbowConn"                                                '// ���R�l�N�^�̕␳
            Call psAdjustElbowConnector
        Case "AdjShapeRoundRect"                                                '// �l�p�`�̊p�ۂݕ␳
            Call psAdjustRoundRect
        Case "AdjShapeBlockArrow"                                               '// �u���b�N���̌X���␳
            Call psAdjustBlockArrowHead
        Case "AdjShapeLine"                                                     '// �����̌X���␳�i0,45,90�x�j
            Call psAdjustLine
        Case "AdjShapeUngroup"                                                  '// �ċA�ŃO���[�v����
            Call psAdjustUngroup
    End Select
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// ���\�b�h�F   ���R�l�N�^�␳
'// �����F       �g�[�i�����g�\�̌��R�l�N�^�̕␳�ʒu�����킹��
'// ////////////////////////////////////////////////////////////////////////////
Private Sub psAdjustElbowConnector()
On Error GoTo ErrorHandler
    Dim topObjName  As String   '// �g�[�i�����g�̒���I�u�W�F�N�g��
    Dim target      As Double   '// �S�R�l�N�^��Adjustment(1)�����̃^�[�Q�b�g�ɍ��킹��B�u�R�l�N�^���~Adjust�l�̍ŏ��l)�v����I�u�W�F�N�g�ɍł��߂��l���̗p����B
    Dim idx         As Integer
    Dim elbows()    As Shape    '// ���R�l�N�^�݂̂��i�[����z��
    Dim cntElbow    As Integer
    Dim bff         As Double
    
    '// ���O�����@//////////
    '// ���O�`�F�b�N�i�A�N�e�B�u�V�[�g�ی�A�I���^�C�v���V�F�C�v�j
    If Not gfPreCheck(protectCont:=True, selType:=TYPE_SHAPE) Then
        Exit Sub
    End If
    
    '// ���R�l�N�^���擾
    cntElbow = 0
    For idx = 1 To ActiveWindow.Selection.ShapeRange.Count         '// shaperange�̊J�n�C���f�b�N�X�͂P����
        If ActiveWindow.Selection.ShapeRange(idx).Connector = msoTrue Then     '// ConnectorFormat�͎��O�ɎQ�Ɖ\���s���Ȃ��߁AIf������l�X�g
            If ActiveWindow.Selection.ShapeRange(idx).ConnectorFormat.Type = msoConnectorElbow Then
                ReDim Preserve elbows(cntElbow)
                Set elbows(cntElbow) = ActiveWindow.Selection.ShapeRange(idx)
                cntElbow = cntElbow + 1
            End If
        End If
    Next
    
    '// �Œ�Q�ȏ�̃R�l�N�^���K�v�B�Ȃ��ꍇ�̓G���[
    If cntElbow < 2 Then
        Call MsgBox(MSG_SHAPE_MULTI_SELECT, vbOKCancel, APP_TITLE)
        Exit Sub
    End If
    
    '// �ŏ���2�̃R�l�N�^�̘A���I�u�W�F�N�g���r���A�g�[�i�����g�̒��_�I�u�W�F�N�g�����擾
    If elbows(0).ConnectorFormat.BeginConnectedShape.Name = elbows(1).ConnectorFormat.BeginConnectedShape.Name Or _
        elbows(0).ConnectorFormat.BeginConnectedShape.Name = elbows(1).ConnectorFormat.EndConnectedShape.Name Then
        topObjName = elbows(0).ConnectorFormat.BeginConnectedShape.Name
    Else
        topObjName = elbows(0).ConnectorFormat.EndConnectedShape.Name
    End If
    
    '// �^�[�Q�b�g�l(�R�l�N�^���~Adjust�l�̍ŏ��l)���擾�@//////////
    target = 0
    For idx = 0 To UBound(elbows)
        With elbows(idx)
            If .ConnectorFormat.BeginConnectedShape.Name = topObjName Then
                bff = .Width * .Adjustments.Item(1)
            Else
                bff = .Width * (1 - .Adjustments.Item(1))
            End If
            
            If target = 0 Then
                target = bff
            ElseIf target > bff Then
                target = bff
            End If
        End With
    Next
    target = Application.WorksheetFunction.Ceiling(target, 0.75)
    
    '// �ŏ��l�ɍ��킹�ăR�l�N�^��ݒ�
    For idx = 0 To UBound(elbows)
        With elbows(idx)
            If .ConnectorFormat.BeginConnectedShape.Name = topObjName Then
                .Adjustments.Item(1) = target / .Width
            Else
                .Adjustments.Item(1) = 1 - (target / .Width)
            End If
        End With
    Next
    
    Exit Sub
ErrorHandler:
    '//
    
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// ���\�b�h�F   �u���b�N���̐�[�p�x�␳
'// �����F       �u���b�N���̐�[�p���A�ł��݊p�Ȃ��̂ɍ��킹��
'// ////////////////////////////////////////////////////////////////////////////
Private Sub psAdjustBlockArrowHead()
    Dim target      As Double   '// �S�u���b�N����Adjustment(1)�����̃^�[�Q�b�g�ɍ��킹��B�u�Z�Ӂ~Adjust�l�̍ŏ��l)�v
    Dim bff         As Double
    Dim idx         As Integer
    
    '// ���O�`�F�b�N�i�A�N�e�B�u�V�[�g�ی�A�I���^�C�v���V�F�C�v�j
    If Not gfPreCheck(protectCont:=True, selType:=TYPE_SHAPE) Then
        Exit Sub
    End If
    
    '// �^�[�Q�b�g�l(�Z�Ӂ~Adjust�l�̍ŏ��l)���擾�@//////////
    target = 0
    For idx = 1 To ActiveWindow.Selection.ShapeRange.Count  '// shaperange�̊J�n�C���f�b�N�X�͂P����
        With ActiveWindow.Selection.ShapeRange(idx)
            If .AutoShapeType = msoShapePentagon Or _
                .AutoShapeType = msoShapeChevron Then
                bff = WorksheetFunction.Min(.Height, .Width) * .Adjustments.Item(1)
                If target = 0 Then
                    target = bff
                ElseIf target > bff Then
                    target = bff
                End If
            End If
        End With
    Next
    
    '// �ŏ��l�ɍ��킹�ău���b�N���̖���ݒ�
    For idx = 1 To ActiveWindow.Selection.ShapeRange.Count  '// shaperange�̊J�n�C���f�b�N�X�͂P����
        With ActiveWindow.Selection.ShapeRange(idx)
            If .AutoShapeType = msoShapePentagon Or _
                .AutoShapeType = msoShapeChevron Then
                .Adjustments.Item(1) = target / WorksheetFunction.Min(.Height, .Width)
            End If
        End With
    Next
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// ���\�b�h�F   �p�̊ۂ��l�p�` �ۂݕ␳
'// �����F       �p�̊ۂ��l�p�`�̊ۂ݂��A�ł�R�i�a�j�̏��������̂ɍ��킹��
'// ////////////////////////////////////////////////////////////////////////////
Private Sub psAdjustRoundRect()
    Dim target      As Double   '// �S�u���b�N����Adjustment(1)�����̃^�[�Q�b�g�ɍ��킹��B�u�Z�Ӂ~Adjust�l�̍ŏ��l)�v
    Dim bff         As Double
    Dim idx         As Integer
    
    '// ���O�`�F�b�N�i�A�N�e�B�u�V�[�g�ی�A�I���^�C�v���V�F�C�v�j
    If Not gfPreCheck(protectCont:=True, selType:=TYPE_SHAPE) Then
        Exit Sub
    End If
    
    '// �^�[�Q�b�g�l(�Z�Ӂ~Adjust�l�̍ŏ��l)���擾 //////////
    target = 0
    For idx = 1 To ActiveWindow.Selection.ShapeRange.Count  '// shaperange�̊J�n�C���f�b�N�X�͂P����
        With ActiveWindow.Selection.ShapeRange(idx)
            If .AutoShapeType = msoShapeRoundedRectangle Then
                bff = WorksheetFunction.Min(.Height, .Width) * .Adjustments.Item(1)
                If target = 0 Then
                    target = bff
                ElseIf target > bff Then
                    target = bff
                End If
            End If
        End With
    Next
    
    '// �ŏ��l�ɍ��킹�Ďl�p�`�̊p��ݒ�
    For idx = 1 To ActiveWindow.Selection.ShapeRange.Count  '// shaperange�̊J�n�C���f�b�N�X�͂P����
        With ActiveWindow.Selection.ShapeRange(idx)
            If .AutoShapeType = msoShapeRoundedRectangle Then
                .Adjustments.Item(1) = target / WorksheetFunction.Min(.Height, .Width)
            End If
        End With
    Next
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// ���\�b�h�F   ���� �p�x�␳
'// �����F       �����̊p�x���A0,45,90�ɕ␳����B�N�_�����ƂɈʒu��␳����
'// ////////////////////////////////////////////////////////////////////////////
Private Sub psAdjustLine()
    Dim lineLen     As Double       '// �I���W�i���̒���
    Dim lineAgl     As Double       '// �I���W�i���̊p�x
    Dim targetAgl   As Double       '// �^�[�Q�b�g�Ƃ���p�x
    Dim idx         As Integer
    Dim bff         As Double
    
    '// ���O�`�F�b�N�i�A�N�e�B�u�V�[�g�ی�A�I���^�C�v���V�F�C�v�j
    If Not gfPreCheck(protectCont:=True, selType:=TYPE_SHAPE) Then
        Exit Sub
    End If
    
    '// �p�x�ݒ�
    For idx = 1 To ActiveWindow.Selection.ShapeRange.Count  '// shaperange�̊J�n�C���f�b�N�X�͂P����
        With ActiveWindow.Selection.ShapeRange(idx)
            If .Type = msoLine Then
                If .Width * .Height <> 0 Then
                    '// �������擾
                    lineLen = Sqr(.Width ^ 2 + .Height ^ 2)
                    '// �p�x���擾
                    lineAgl = WorksheetFunction.Degrees(Atn((.Height) / (.Width)))
                    Select Case lineAgl
                        Case Is >= 70   '// 90�x�ɕ␳
                            bff = .Width
                            .Width = 0
                            If .HorizontalFlip Then
                                .Left = .Left + bff
                            End If
                        Case Is <= 30
                            bff = .Height
                            .Height = 0
                            If .VerticalFlip Then
                                .Top = .Top + bff
                            End If
                        Case Else   '// 45�x�ɕ␳
                            .Height = Sqr(lineLen ^ 2 / 2)
                            .Width = .Height
                    End Select
                End If
            End If
        End With
    Next
'Debug.Print "len: " & lineLen
'Debug.Print targetAgl & " / " & lineAgl

End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// ���\�b�h�F   �ċA�ŃO���[�v����
'// �����F       �l�X�g�����O���[�v�����ׂĉ�������B�O���[�v�������� _sub�Ɏ���
'// ////////////////////////////////////////////////////////////////////////////
Private Sub psAdjustUngroup()
    Dim idx         As Integer
    Dim sh          As Shape
    
    '// ���O�`�F�b�N�i�A�N�e�B�u�V�[�g�ی�A�I���^�C�v���V�F�C�v�j
    If Not gfPreCheck(protectCont:=True, selType:=TYPE_SHAPE) Then
        Exit Sub
    End If
    
    For idx = 1 To ActiveWindow.Selection.ShapeRange.Count  '// shaperange�̊J�n�C���f�b�N�X�͂P����
        Call psAdjustUngroup_sub(ActiveWindow.Selection.ShapeRange(idx))
    Next

End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// ���\�b�h�F   �ċA�ŃO���[�v����
'// �����F       �O���[�v����������
'// ////////////////////////////////////////////////////////////////////////////
Private Sub psAdjustUngroup_sub(targetShape As Shape)
    Dim sh As Shape
    
    If targetShape.Type = msoGroup Then
        For Each sh In targetShape.Ungroup
            Call psAdjustUngroup_sub(sh)
        Next
    End If
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// END
'// ////////////////////////////////////////////////////////////////////////////
