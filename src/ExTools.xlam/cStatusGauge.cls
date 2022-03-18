VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cStatusGauge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'// ////////////////////////////////////////////////////////////////////////////
'// �v���W�F�N�g   : �G�N�Z���g���c�[��
'// �^�C�g��       : �X�e�[�^�X�o�[ �e�L�X�g�Q�[�W�\���N���X
'//
'// ���W���[��     : cStatusGauge
'// �ŏI�C����     : 2009/10/09
'// �ŏI�C����     : ��������
'//
'// ����           : �X�e�[�^�X�o�[�Ƀe�L�X�g�Ńv���O���X�o�[��\������
'// ////////////////////////////////////////////////////////////////////////////
'// Copyright (c) by Koichiro.  All Rights Reserved.
'// ////////////////////////////////////////////////////////////////////////////
Option Explicit
Option Base 0


'// ////////////////////////////////////////////////////////////////////////////
'// �v���C�x�[�g�ϐ�
Private pMaxVal                       As Double       '// �v���p�e�B�l�ێ��F�ő�l
Private pCurrentVal                   As Double       '// �v���p�e�B�l�ێ��F���ݒl
Private pPercentage                   As Integer      '// �v���p�e�B�l�ێ��F�p�[�Z���g


'///////////////////////////////////////////////////////////////////////////////
'// �R���X�g���N�^
Private Sub Class_Initialize()
  Application.StatusBar = False
  pMaxVal = 100
  pCurrentVal = 0
  pPercentage = 0
End Sub


'///////////////////////////////////////////////////////////////////////////////
'// �f�X�g���N�^
Private Sub Class_Terminate()
  Application.StatusBar = False
End Sub


'///////////////////////////////////////////////////////////////////////////////
'// �v���p�e�B: �ő�l
Public Property Let MaxVal(newValue As Double)
  pMaxVal = newValue
  Call psPutStatusGauge
End Property

Public Property Get MaxVal() As Double
  MaxVal = pMaxVal
End Property


'///////////////////////////////////////////////////////////////////////////////
'// �v���p�e�B: ���ݒl
Public Property Let CurrentVal(newValue As Double)
  pCurrentVal = newValue
  Call psPutStatusGauge
End Property

Public Property Get CurrentVal() As Double
  CurrentVal = pMaxVal
End Property


'///////////////////////////////////////////////////////////////////////////////
'// �v���p�e�B: �p�[�Z���e�[�W (ReadOnly)
Public Property Get Percentage() As Integer
  Percentage = pPercentage
End Property


'// ////////////////////////////////////////////////////////////////////////////
'// ���\�b�h�F   �l�̑���
'// �����F       �����̒l�𑝕��Ƃ��Ď󂯎��A�Q�[�W�ɔ��f������B
'// �����F       val2add: ����
'///////////////////////////////////////////////////////////////////////////////
Public Function addValue(value2add As Double)
  pCurrentVal = pCurrentVal + value2add
  Call psPutStatusGauge
End Function


'// ////////////////////////////////////////////////////////////////////////////
'// ���\�b�h�F   �Q�[�W�\��
'// �����F       �X�e�[�^�X�o�[�ɃQ�[�W��\������B
'// ////////////////////////////////////////////////////////////////////////////
Private Sub psPutStatusGauge()
  Dim newVal  As Integer
  
  newVal = CInt(((pCurrentVal / pMaxVal) * 100) - 0.5)
  If newVal <> pPercentage Then
    pPercentage = newVal
    Application.StatusBar = CStr(pPercentage) & "%  " & String(CInt((newVal - 0.5) / 2), "|")
  End If
End Sub


'// ////////////////////////////////////////////////////////////////////////////
'// END
'// ////////////////////////////////////////////////////////////////////////////
