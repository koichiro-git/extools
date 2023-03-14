Attribute VB_Name = "mdlBarcode"
'// ////////////////////////////////////////////////////////////////////////////
'// �v���W�F�N�g   : �g���c�[�� �ǉ��p�b�N
'// �^�C�g��       : QR�R�[�h�\��
'// ���W���[��     : mdlBarcode
'// ����           : �Z�����e����QR�R�[�h�摜�𐶐�����B
'//                  Access�p�����^�C����Excel�ł͓���ۏ؂���Ă��Ȃ����߁A�t���[��Web�T�[�r�X���g�p
'// ////////////////////////////////////////////////////////////////////////////
'// Copyright (c) by Koichiro.
'// ////////////////////////////////////////////////////////////////////////////
Option Explicit
Option Base 0

'// �Â��o�[�W�����ł��R���p�C����ʂ����߃o�[�R�[�h�̎Q�Ɛݒ�͍s��Ȃ��B���̂��ߒ萔�����߂Ē�`
'//https://rdr.utopiat.net/�v

'Enum BarcodeStyle
'    JAN13 = 2
'    JAN8 = 3
'    Casecode = 4
'    NW7 = 5
'    Code39 = 6
'    Code128 = 7
'    QR = 11
'End Enum


'// ////////////////////////////////////////////////////////////////////////////
'// ���\�b�h�F   ���{���{�^���R�[���o�b�N�Ǘ�
'// �����F       ���{������̃R�[���o�b�N�������ǂ�
'//              �����ꂽ�R���g���[����ID����ɏ������Ăяo���B
'// �����F       control �ΏۃR���g���[��
'// ////////////////////////////////////////////////////////////////////////////
Public Sub ribbonCallback_BarCode(control As IRibbonControl)
    Select Case control.ID
        Case "QRCode"                       '// QR�R�[�h
            Call psDrawBarCode
    End Select
End Sub



'// ////////////////////////////////////////////////////////////////////////////
'// ���\�b�h�F   �o�[�R�[�h�`��
'// �����F       �I�����ꂽ�Z���̒l�����Ƀo�[�R�[�h��`�悷��
'// ////////////////////////////////////////////////////////////////////////////
Sub psDrawBarCode()
Attribute psDrawBarCode.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim tCell     As Range    '// �ϊ��ΏۃZ��
    
    '//
'    If Application.Version < 16 Then
'        Call MsgBox("�o�[�R�[�h��Excel2016�ȍ~�̃o�[�W�����ł̂ݎg�p�\�ł�")
'        Exit Sub
'    End If
    
    '// ���O�`�F�b�N�i�A�N�e�B�u�V�[�g�ی�A�I���^�C�v���Z���j
    If Not gfPreCheck(protectCont:=True, selType:=TYPE_RANGE) Then
        Exit Sub
    End If
    
    Call gsSuppressAppEvents
    
    If Selection.Count > 1 Then
        For Each tCell In Selection.SpecialCells(xlCellTypeConstants, xlNumbers + xlTextValues)
            Call psDrawBarCode_sub2(tCell)
        Next
    Else
        Call psDrawBarCode_sub2(ActiveCell)
    End If
    
    Call gsResumeAppEvents
    Exit Sub
    
ErrorHandler:
    '//
    Call gsResumeAppEvents
End Sub
    
    
Private Sub psDrawBarCode_sub(tCell As Range)
    Dim obj         As OLEObject    '// �o�[�R�[�h��ActiveX�I�u�W�F�N�g�{�̂��i�[
    Dim bcd         As Object       '// ActiveX�I�u�W�F�N�g�����̃o�[�R�[�h�`��I�u�W�F�N�g���i�[
    Dim pctCamera   As Object       '// �摜�ɕϊ�����ۂɈꎞ�I��CopyPicture�ɉ摜�Ƃ��ĔF�������邽�߂̃J�����I�u�W�F�N�g
        
    Set obj = ActiveSheet.OLEObjects.Add(ClassType:="BARCODE.BarCodeCtrl.1", Link:=False, DisplayAsIcon:=False)
    
    '// �ݒ�(�v���p�e�B�y�[�W��)
    Set bcd = obj.Object
    With bcd
        .style = 11
        .Validation = 1
'                .Refresh
    End With
    
    obj.LinkedCell = tCell.Address      '// Linked Cell ��Style����ɍŌ�ɐݒ肵�Ȃ��ƁA�Z���l��������
    obj.Top = tCell.Top                 '// �Ō�ɃT�C�Y�ύX���邱�Ƃŕ`�惊�t���b�V��
    obj.Left = tCell.Left
    obj.Width = tCell.Width
    obj.Height = tCell.Height
    
    '// �J�����Ƃ��ĕ����iCopyPicture�ł�ActiveX�I�u�W�F�N�g�͔F������Ȃ����ߕK�v�ȏ����j
    tCell.Select
    tCell.Copy
    Set pctCamera = ActiveSheet.Pictures.Paste(Link:=True)
    '// �Z�����摜�Ƃ��ăR�s�[�E�\��t������B
    Call tCell.CopyPicture(xlPrinter, xlPicture)
    Call ActiveSheet.Paste
    '// �J�������폜
    pctCamera.Delete
    '// ActiveX�I�u�W�F�N�g���폜
    obj.Delete
End Sub

    
    
Private Sub psDrawBarCode_sub2(tCell As Range)
    Dim obj         As Shape    '// �摜�\�t�p�V�F�C�v
    
    Set obj = ActiveSheet.Shapes.AddPicture("http://api.qrserver.com/v1/create-qr-code/?data=" & tCell.Text & _
                                            "!&size=300x300", False, True, tCell.Left, tCell.Top, tCell.Width, tCell.Height)
'    obj.Top = tCell.Top
'    obj.Left = tCell.Left
'    obj.Width = tCell.Width
'    obj.Height = tCell.Height
    
End Sub
    
    
'// �J�������g���ăR�s�[
'    Selection.Copy
'    Range("M13").Select
'    ActiveSheet.Pictures.Paste(Link:=True).Select
'    ActiveSheet.Shapes.Range(Array("Picture 60")).Select
'    Application.CutCopyMode = False
'    Selection.Copy
'    Range("M14").Select
'    ActiveSheet.Pictures.Paste.Select
    
    
    
    
'
'
'
'
'    For idxRow = 7 To 12
'
'        Set obj = ActiveSheet.OLEObjects.Add(ClassType:="BARCODE.BarCodeCtrl.1", Link:=False, DisplayAsIcon:=False)
'        obj.Top = Cells(idxRow, 2).Top
'        obj.Left = Cells(idxRow, 2).Left
'        obj.Width = Cells(idxRow, 2).Width
'        obj.Height = Cells(idxRow, 2).Height
'
' '       Set bcd = CreateObject("BARCODE.BarCodeCtrl.1")
'        Set bcd = obj.Object
'
'        With bcd
'          .Style = BarcodeStyle.Code128
''          .LinkedCell = Cells(idxRow, 2).Address
''          .SubStyle = BC_Substyle
'          .Validation = 1
''          .LineWeight = BC_LineWeight
''          .Direction = BC_Direction
''          .ShowData = BC_ShowData
''          .ForeColor = BC_ForeColor
''          .BackColor = BC_BackColor
'          .Refresh
'         End With
'        'Linked Cell �͍Ō�ɐݒ肵�Ȃ��ƁA�Z���l��������B�B�B
'        obj.LinkedCell = Cells(idxRow, 2).Address
'
'    Next
'
    'JAN,CODE39,ITF,NW-7,CODE128�ȂǁA
    'JAN/EAN/UPC ITF CODE39 NW-7(CODABAR) CODE128 ��\5���
    



' '�v���p�e�B�ɂ��Ă͈ȉ�URL��MSDN�Q��
'    'https://msdn.microsoft.com/ja-jp/library/cc427149.aspx
'
'    Const BC_Style As Integer = 7
'    '�X�^�C��
'    '0: UPC-A, 1: UPC-E, 2: JAN-13, 3: JAN-8, 4: Casecode, 5: NW-7,
'    '6: Code-39, 7: Code-128, 8: U.S. Postnet, 9: U.S. Postal FIM, 10: �X�֕��̕\���p�r�i���{�j
'
'    Const BC_Substyle As Integer = 0
'    '�T�u�X�^�C�� (���LURL�Q��)
'    'https://msdn.microsoft.com/ja-jp/library/cc427156.aspx
'
'    Const BC_Validation As Integer = 1
'    '�f�[�^�̊m�F
'    '0: �m�F����, 1: �����Ȃ�v�Z��␳, 2: �����Ȃ��\��
'    'Code39/NW-7�̏ꍇ�A�u1�v�ŃX�^�[�g/�X�g�b�v����(*)�������I�ɒǉ�
'
'    Const BC_LineWeight As Integer = 3
'    '���̑���
'    '0: �ɍא�, 1:�א�, 2:���א�, 3:�W��, 4:������, 5: ����, 6:�ɑ���, 7:���ɑ���
'
'    Const BC_Direction As Integer = 0
'    '�o�[�R�[�h�̕\������
'    '0: 0�x, 1: 90�x, 2: 180�x, 3: 270�x�@[0]���W��
'
'    Const BC_ShowData As Integer = 1
'    '�f�[�^�̕\��
'    '0: �\������, 1:�\���L��
'
'    Const BC_ForeColor As Long = rgbBlack
'    '�O�i�F�̎w��
'
'    Const BC_BackColor As Long = rgbWhite
'    '�w�i�F�̎w��
'
'    'rgbBlack�Ȃǂ̐F�萔�͈ȉ�URL��MSDN�Q��
'    'https://msdn.microsoft.com/ja-jp/VBA/Excel-VBA/articles/xlrgbcolor-enumeration-excel
'
'   '**�o�[�R�[�h���̏�������

'End Sub


'// https://translate.google.pl/?sl=en&tl=ja&text=hello&op=translate



