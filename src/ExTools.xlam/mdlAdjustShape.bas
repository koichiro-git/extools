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
        Case "AdjShapeOrderTile"                                                '// �O���b�h�ɐ���
            Call psDistributeShapeGrid
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
'// ���\�b�h�F   �O���b�h����
'// �����F       ���C������
'// ////////////////////////////////////////////////////////////////////////////
Private Sub psDistributeShapeGrid()
On Error GoTo ErrorHandler
    Dim tls             As Shape    '// Top-Left-Shape. ����̊�Ƃ���V�F�C�v
    Dim allShapes()     As Shape    '// ���ׂẴV�F�C�v���i�[
    Dim rowHeader()     As Shape    '// �s�w�b�_�i�c���j�̃V�F�C�v���i�[
    Dim colHeader()     As Shape    '// ��w�b�_�i�����j�̃V�F�C�v���i�[
    
    '// �S�V�F�C�v��z��Ɋi�[
    allShapes = pfGetAllShapes(Selection.ShapeRange)
    '// TopLeft���擾
    Set tls = pfGetTopLeftObject(Selection.ShapeRange)
    
    '// �s�w�b�_�ɂ�����V�F�C�v�̔z���ݒ�
    rowHeader = pfGetRowHeader(tls, allShapes)
    colHeader = pfGetColHeader(tls, allShapes)
    
    Call psAdjustAllShapes(allShapes, rowHeader, colHeader)
    
    '// ��n��
    Call tls.Select
    Exit Sub
    
ErrorHandler:
    Call gsResumeAppEvents
    Call gsShowErrorMsgDlg("psDistributeShapeGrid", Err)
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
rslt.Fill.ForeColor.ObjectThemeColor = msoThemeColorAccent2
End Function


'// ////////////////////////////////////////////////////////////////////////////
'// ���\�b�h�F   �O���b�h����
'// �����F       �s�w�b�_�i�c���j�擾
'// ////////////////////////////////////////////////////////////////////////////
Private Function pfGetRowHeader(tls As Shape, ary() As Shape) As Shape()
On Error GoTo ErrorHandler
'    Dim shp         As Shape
    Dim rslt()      As Shape
    Dim i           As Integer
    Dim bff         As Shape
    Dim idxS1       As Long
    Dim idxS2       As Long

    '// �c���ɊY������I�u�W�F�N�g��z��Ɋi�[
    ReDim rslt(0)
    For i = 0 To UBound(ary)
        If ary(i).Left < (tls.Left + tls.Width) Then
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
    
    '// �ʒu�␳(�I������)
    Range("A1").Select
    
    For i = 0 To UBound(rslt)
        rslt(i).TextFrame2.TextRange.Characters.Text = rslt(i).TextFrame2.TextRange.Characters.Text & " �c�� head" & i
'        rslt(i).Left = tls.Left
        Call rslt(i).Select(Replace:=False)
    Next
    
    If UBound(rslt) > 1 Then    '// ����iDistribute�j�͂R�ȏ�̃I�u�W�F�N�g�������ƃG���[�ɂȂ邽��
        Call Selection.ShapeRange.Distribute(msoDistributeVertically, False)
    End If
    
    pfGetRowHeader = rslt
    Exit Function
    
ErrorHandler:
    Call gsShowErrorMsgDlg("pfGetRowHeader", Err)
End Function


'// ////////////////////////////////////////////////////////////////////////////
'// ���\�b�h�F   �O���b�h����
'// �����F       ��w�b�_�i�����j�擾
'// ////////////////////////////////////////////////////////////////////////////
Private Function pfGetColHeader(tls As Shape, ary() As Shape) As Shape()
On Error GoTo ErrorHandler
'    Dim shp         As Shape
    Dim rslt()      As Shape
    Dim i           As Integer
    Dim bff         As Shape
    Dim idxS1      As Long
    Dim idxS2      As Long
    
    '// �����ɊY������I�u�W�F�N�g��z��Ɋi�[
    ReDim rslt(0)
    For i = 0 To UBound(ary)
        If ary(i).Top < (tls.Top + tls.Height) Then
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
    
    '// �ʒu�␳(�I������)
    Range("A1").Select

    For i = 0 To UBound(rslt)
        rslt(i).TextFrame2.TextRange.Characters.Text = rslt(i).TextFrame2.TextRange.Characters.Text & " ���� head" & i
'        rslt(i).Top = tls.Top
        Call rslt(i).Select(Replace:=False)
    Next
    
    If UBound(rslt) > 1 Then   '// ����iDistribute�j�͂R�ȏ�̃I�u�W�F�N�g�������ƃG���[�ɂȂ邽��
        Call Selection.ShapeRange.Distribute(msoDistributeHorizontally, False)
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
    Dim idxHead           As Integer
    
    '// �S�V�F�C�v�ł̃��[�v
    For idx = 0 To UBound(allShapes)
        '// �s�w�b�_�i�c���j�ł̃��[�v
        For idxHead = 0 To UBound(rowHeader)
            If allShapes(idx).Top < rowHeader(idxHead).Top + rowHeader(idxHead).Height Then
                allShapes(idx).Top = rowHeader(idxHead).Top
                allShapes(idx).Height = rowHeader(idxHead).Height
                Exit For
            End If
        Next
        
        '// ��w�b�_�i�����j�ł̃��[�v
        For idxHead = 0 To UBound(colHeader)
            If allShapes(idx).Left < colHeader(idxHead).Left + colHeader(idxHead).Width Then
                allShapes(idx).Left = colHeader(idxHead).Left
                allShapes(idx).Width = colHeader(idxHead).Width
                Exit For
            End If
        Next
    Next
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// END
'// ////////////////////////////////////////////////////////////////////////////
