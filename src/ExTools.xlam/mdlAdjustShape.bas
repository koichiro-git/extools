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
'// �R���p�C���X�C�b�`�i"EXCEL" / "POWERPOINT"�j
#Const OFFICE_APP = "EXCEL"


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
        Case "AdjShapeOrderTile"                                                '// �O���b�h�ɐ���
            Call psDistributeShapeGrid(0)
        Case "AdjShapeOrderTile_1"                                                '// �O���b�h�ɐ��� 1pt
            Call psDistributeShapeGrid(1)
        Case "AdjShapeOrderTile_2"                                                '// �O���b�h�ɐ��� 2pt
            Call psDistributeShapeGrid(2)
        Case "AdjShapeOrderTile_3"                                                '// �O���b�h�ɐ��� 3pt
            Call psDistributeShapeGrid(3)
        Case "AdjShapeOrderTile_4"                                                '// �O���b�h�ɐ��� 4pt
            Call psDistributeShapeGrid(4)
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
    
    '// ���O���� //////////
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
    target = gfCeiling(target, 0.75)
    
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
                bff = gfMin2(.Height, .Width) * .Adjustments.Item(1)
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
                .Adjustments.Item(1) = target / gfMin2(.Height, .Width)
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
                bff = gfMin2(.Height, .Width) * .Adjustments.Item(1)
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
                .Adjustments.Item(1) = target / gfMin2(.Height, .Width)
            End If
        End With
    Next
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// ���\�b�h�F   ���� �p�x�␳
'// �����F       �����̊p�x���A0,45,90�x�ɕ␳����B���̈ʒu�̒��S�����]������
'// ////////////////////////////////////////////////////////////////////////////
Private Sub psAdjustLine()
    Dim idx         As Integer
    
    '// ���O�`�F�b�N�i�A�N�e�B�u�V�[�g�ی�A�I���^�C�v���V�F�C�v�j
    If Not gfPreCheck(protectCont:=True, selType:=TYPE_SHAPE) Then
        Exit Sub
    End If
    
    '// �p�x�ݒ�
    For idx = 1 To ActiveWindow.Selection.ShapeRange.Count  '// shaperange�̊J�n�C���f�b�N�X�͂P����
        With ActiveWindow.Selection.ShapeRange(idx)
            If .Type = msoLine Then
                If .Width * .Height <> 0 Then
                    Select Case Atn(.Height / .Width) * 180 / (Atn(1) * 4)
                        Case Is <= 30   '// 0�x�ɕ␳
                            .Top = IIf(.VerticalFlip, .Top - .Height / 2, .Top + .Height / 2)
                            .Height = 0
                        Case Is >= 70   '// 90�x�ɕ␳
                            .Left = IIf(.VerticalFlip, .Left - .Width / 2, .Left + .Width / 2)
                            .Width = 0
                        Case Else   '// 45�x�ɕ␳
                            If .Height > .Width Then
                                .Left = .Left - (.Height - .Width) / 2
                                .Width = .Height
                            Else
                                .Top = .Top - (.Width - .Height) / 2
                                .Height = .Width
                            End If
                    End Select
                End If
            End If
        End With
    Next
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// ���\�b�h�F   �ċA�ŃO���[�v����
'// �����F       �l�X�g�����O���[�v�����ׂĉ�������B�O���[�v�������� _sub�Ɏ���
'// ////////////////////////////////////////////////////////////////////////////
Private Sub psAdjustUngroup()
On Error GoTo ErrorHandler
    Dim idx         As Integer
'    Dim sh          As Shape
    
    '// ���O�`�F�b�N�i�A�N�e�B�u�V�[�g�ی�A�I���^�C�v���V�F�C�v�j
    If Not gfPreCheck(protectCont:=True, selType:=TYPE_SHAPE) Then
        Exit Sub
    End If
    
    For idx = 1 To ActiveWindow.Selection.ShapeRange.Count  '// shaperange�̊J�n�C���f�b�N�X�͂P����
        Call psAdjustUngroup_sub(ActiveWindow.Selection.ShapeRange(idx))
    Next
    Exit Sub

ErrorHandler:
#If OFFICE_APP = "EXCEL" Then
    Call gsResumeAppEvents
#End If
    Call gsShowErrorMsgDlg("psAdjustUngroup", Err)
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
'// ���\�b�h�F   �O���b�h����
'// �����F       ���C������
'// ////////////////////////////////////////////////////////////////////////////
Private Sub psDistributeShapeGrid(spacing As Integer)
'On Error GoTo ErrorHandler
    Dim tls             As Shape    '// Top-Left-Shape. ����̊�Ƃ���V�F�C�v
    Dim allShapes()     As Shape    '// ���ׂẴV�F�C�v���i�[
    Dim rowHeader()     As Shape    '// �s�w�b�_�i�c���j�̃V�F�C�v���i�[
    Dim colHeader()     As Shape    '// ��w�b�_�i�����j�̃V�F�C�v���i�[
    
    '// ���O�`�F�b�N�i�A�N�e�B�u�V�[�g�ی�A�I���^�C�v���V�F�C�v�j
    If Not gfPreCheck(protectCont:=True, selType:=TYPE_SHAPE) Then
        Exit Sub
    End If
    
#If OFFICE_APP = "EXCEL" Then
    allShapes = pfGetAllShapes(Selection.ShapeRange)                '// �S�V�F�C�v��z��Ɋi�[
    Set tls = pfGetTopLeftObject(Selection.ShapeRange)              '// TopLeft���擾
#ElseIf OFFICE_APP = "POWERPOINT" Then
    allShapes = pfGetAllShapes(ActiveWindow.Selection.ShapeRange)   '// �S�V�F�C�v��z��Ɋi�[
    Set tls = pfGetTopLeftObject(ActiveWindow.Selection.ShapeRange) '// TopLeft���擾
#End If
    
    '// �s�w�b�_�ɂ�����V�F�C�v�̔z���ݒ�
    rowHeader = pfGetRowHeader(tls, allShapes, spacing)
    colHeader = pfGetColHeader(tls, allShapes, spacing)
    
    Call psAdjustAllShapes(allShapes, rowHeader, colHeader)
    Exit Sub
    
ErrorHandler:
#If OFFICE_APP = "EXCEL" Then
    Call gsResumeAppEvents
#End If
'    Call gsShowErrorMsgDlg("psDistributeShapeGrid", Err)
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// ���\�b�h�F   �O���b�h����
'// �����F       �I�����ꂽ�V�F�C�v��S�Ĕz��Ɋi�[
'// ////////////////////////////////////////////////////////////////////////////
Public Function pfGetAllShapes(rng As ShapeRange) As Shape()
    Dim shp         As Shape
    Dim rslt()      As Shape
    Dim i           As Integer
    
    ReDim rslt(0)
    For Each shp In rng
        If rslt(0) Is Nothing Then
            Set rslt(0) = shp
        Else
            ReDim Preserve rslt(UBound(rslt) + 1)
            Set rslt(UBound(rslt)) = shp
        End If
    Next
    
    pfGetAllShapes = rslt
End Function


'// ////////////////////////////////////////////////////////////////////////////
'// ���\�b�h�F   �O���b�h����
'// �����F       ��ƂȂ�ATopLeft�ʒu�̃V�F�C�v���擾
'// ////////////////////////////////////////////////////////////////////////////
Public Function pfGetTopLeftObject(rng As ShapeRange) As Shape
    Dim shp         As Shape
    Dim rslt        As Shape
    
    Set rslt = rng(1)
    
    '// Top���ł��������V�F�C�v���擾
    For Each shp In rng
        If shp.Top < rslt.Top Then
            Set rslt = shp
        End If
    Next
    
    '// �ŏ�Top�̃V�F�C�v�̉��ӂ���Top���������A���ŏ���Left�����V�F�C�v���擾
    For Each shp In rng
        If shp.Top < (rslt.Top + rslt.Height) And shp.Left < rslt.Left Then
            Set rslt = shp
        End If
    Next
    
    Set pfGetTopLeftObject = rslt
    
'//�@�Ԃɂ���
'rslt.Fill.ForeColor.ObjectThemeColor = msoThemeColorAccent2
End Function


'// ////////////////////////////////////////////////////////////////////////////
'// ���\�b�h�F   �O���b�h����
'// �����F       �s�w�b�_�i�c���j�擾
'// ////////////////////////////////////////////////////////////////////////////
Private Function pfGetRowHeader(tls As Shape, ary() As Shape, spacing As Integer) As Shape()
'On Error GoTo ErrorHandler
'    Dim shp         As Shape
    Dim rslt()      As Shape
    Dim i           As Integer
    Dim bff         As Shape
    Dim idxS1       As Long
    Dim idxS2       As Long
    Dim heightTotal As Long     '// �w�b�_�̃V�F�C�v�̍����̍��v

    '// �c���ɊY������I�u�W�F�N�g��z��Ɋi�[
    ReDim rslt(0)
    For i = 0 To UBound(ary)
        If ary(i).Left < (tls.Left + (tls.Width * 0.9)) Then
            If Not rslt(0) Is Nothing Then
                ReDim Preserve rslt(UBound(rslt) + 1)
            End If
            Set rslt(UBound(rslt)) = ary(i)
        End If
    Next
    
    '// �\�[�g
    idxS1 = 0
    ' �S�e�[�u���̑O����̃��[�v
    Do While idxS1 < UBound(rslt)
        idxS2 = UBound(rslt)
        ' �I�[���猻�݈ʒu��O�܂ł̃��[�v
        Do While idxS2 > idxS1
            ' �����ւ�����
            If rslt(idxS2).Top < rslt(idxS1).Top Then
                ' �����ւ�
                Set bff = rslt(idxS2)
                Set rslt(idxS2) = rslt(idxS1)
                Set rslt(idxS1) = bff
            End If
            ' �O��
            idxS2 = idxS2 - 1
        Loop
        ' ����
        idxS1 = idxS1 + 1
    Loop
    
    '// �I������
#If OFFICE_APP = "EXCEL" Then
    tls.TopLeftCell.Select
#ElseIf OFFICE_APP = "POWERPOINT" Then
    Call ActiveWindow.Selection.Unselect
#End If
    
    For i = 0 To UBound(rslt)
        Call rslt(i).Select(Replace:=False)
        heightTotal = heightTotal + rslt(i).Height
'        rslt(i).Line.ForeColor.RGB = vbRed
    Next
    
    '// �I�u�W�F�N�g���d�Ȃ��Ă���ꍇ�i�����̍��v���Ō�̃I�u�W�F�N�g�̏I�_�����������j�́A�z�u���L����
    If UBound(rslt) > 0 Then
        If heightTotal >= (rslt(UBound(rslt)).Top + rslt(UBound(rslt)).Height) - tls.Top _
              Or spacing > 0 Then
'            rslt(UBound(rslt)).Line.ForeColor.RGB = vbRed
            rslt(UBound(rslt)).Top = tls.Top + heightTotal - rslt(UBound(rslt)).Height + UBound(rslt) * spacing
        End If
    End If
    
    
    If UBound(rslt) > 1 Then    '// ����iDistribute�j�͂R�ȏ�̃I�u�W�F�N�g�������ƃG���[�ɂȂ邽��
#If OFFICE_APP = "EXCEL" Then
        Call Selection.ShapeRange.Distribute(msoDistributeVertically, False)
#ElseIf OFFICE_APP = "POWERPOINT" Then
        Call ActiveWindow.Selection.ShapeRange.Distribute(msoDistributeVertically, False)
#End If
    End If
    
    pfGetRowHeader = rslt
    Exit Function
    
ErrorHandler:
'    Call gsShowErrorMsgDlg("pfGetRowHeader", Err)
End Function


'// ////////////////////////////////////////////////////////////////////////////
'// ���\�b�h�F   �O���b�h����
'// �����F       ��w�b�_�i�����j�擾
'// ////////////////////////////////////////////////////////////////////////////
Private Function pfGetColHeader(tls As Shape, ary() As Shape, spacing As Integer) As Shape()
'On Error GoTo ErrorHandler
'    Dim shp         As Shape
    Dim rslt()      As Shape
    Dim i           As Integer
    Dim bff         As Shape
    Dim idxS1       As Long
    Dim idxS2       As Long
    Dim widthTotal  As Long     '// �w�b�_�̃V�F�C�v�̕��̍��v
    
    '// �����ɊY������I�u�W�F�N�g��z��Ɋi�[
    widthTotal = 0
    ReDim rslt(0)
    For i = 0 To UBound(ary)
        If ary(i).Top < (tls.Top + (tls.Height * 0.9)) Then
            If Not rslt(0) Is Nothing Then
                ReDim Preserve rslt(UBound(rslt) + 1)
            End If
            Set rslt(UBound(rslt)) = ary(i)
        End If
    Next
    
    '// �\�[�g
    idxS1 = 0
    Do While idxS1 < UBound(rslt)                       '// �O����̃��[�v
        idxS2 = UBound(rslt)
        Do While idxS2 > idxS1                          '// �I�[���猻�݈ʒu��O�܂ł̃��[�v
            If rslt(idxS2).Left < rslt(idxS1).Left Then '// �\�[�g����ւ�����
                Set bff = rslt(idxS2)
                Set rslt(idxS2) = rslt(idxS1)
                Set rslt(idxS1) = bff
            End If
            idxS2 = idxS2 - 1
        Loop
        idxS1 = idxS1 + 1
    Loop
    
    '// �I������
#If OFFICE_APP = "EXCEL" Then
    tls.TopLeftCell.Select
#ElseIf OFFICE_APP = "POWERPOINT" Then
    Call ActiveWindow.Selection.Unselect
#End If

    For i = 0 To UBound(rslt)
        Call rslt(i).Select(Replace:=False)
        widthTotal = widthTotal + rslt(i).Width
'        rslt(i).Line.ForeColor.RGB = vbBlue
    Next
    
    '// �I�u�W�F�N�g���d�Ȃ��Ă���ꍇ�i���̍��v���Ō�̃I�u�W�F�N�g�̏I�_�����������j�́A�z�u���L����
    If UBound(rslt) > 0 Then
        If widthTotal >= (rslt(UBound(rslt)).Left + rslt(UBound(rslt)).Width) - tls.Left _
              Or spacing > 0 Then
'            rslt(UBound(rslt)).Line.ForeColor.RGB = vbBlue
            rslt(UBound(rslt)).Left = tls.Left + widthTotal - rslt(UBound(rslt)).Width + UBound(rslt) * spacing
        End If
    End If
    
    If UBound(rslt) > 1 Then    '// ����iDistribute�j�͂R�ȏ�̃I�u�W�F�N�g�������ƃG���[�ɂȂ邽��
#If OFFICE_APP = "EXCEL" Then
        Call Selection.ShapeRange.Distribute(msoDistributeHorizontally, False)
#ElseIf OFFICE_APP = "POWERPOINT" Then
        Call ActiveWindow.Selection.ShapeRange.Distribute(msoDistributeHorizontally, False)
#End If
    End If
    
    pfGetColHeader = rslt
    Exit Function
    
ErrorHandler:
    Call gsShowErrorMsgDlg("pfGetColHeader", Err)
End Function


'// ////////////////////////////////////////////////////////////////////////////
'// �S�V�F�C�v�̔z�u
Private Sub psAdjustAllShapes(allShapes() As Shape, rowHeader() As Shape, colHeader() As Shape)
    Dim idx                 As Integer
    Dim idxHead             As Integer
    Dim bff                 As Double   '// ����ΏۃV�F�C�v�̒����ʒu���i�[
    
    '// �S�V�F�C�v�ł̃��[�v
    For idx = 0 To UBound(allShapes)
        '// �s�w�b�_�i�c���j�ł̃��[�v
        For idxHead = 0 To UBound(rowHeader)
            bff = allShapes(idx).Top + allShapes(idx).Height / 2    '// �ΏۃI�u�W�F�N�g�̒����|�W�V�����i�c�j
'            If bff >= allShapes(idx).Top And bff <= rowHeader(idxHead).Top + rowHeader(idxHead).Height Then
            If bff >= rowHeader(idxHead).Top And bff <= rowHeader(idxHead).Top + rowHeader(idxHead).Height Then
                allShapes(idx).Top = rowHeader(idxHead).Top
                allShapes(idx).Height = rowHeader(idxHead).Height
                Exit For
            End If
        Next
        
        '// ��w�b�_�i�����j�ł̃��[�v
        For idxHead = 0 To UBound(colHeader)
            bff = allShapes(idx).Left + allShapes(idx).Width / 2    '// �ΏۃI�u�W�F�N�g�̒����|�W�V�����i�c�j
'            If bff >= allShapes(idx).Left And bff <= colHeader(idxHead).Left + colHeader(idxHead).Width Then
            If bff >= colHeader(idxHead).Left And bff <= colHeader(idxHead).Left + colHeader(idxHead).Width Then
                allShapes(idx).Left = colHeader(idxHead).Left
                allShapes(idx).Width = colHeader(idxHead).Width
                Exit For
            End If
        Next
        
        '// �I�����������Ƃɖ߂�
        Call allShapes(idx).Select(Replace:=False)
    Next
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// END
'// ////////////////////////////////////////////////////////////////////////////


