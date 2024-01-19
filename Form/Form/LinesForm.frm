VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} LinesForm 
   Caption         =   "LinesForm"
   ClientHeight    =   2430
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4680
   OleObjectBlob   =   "LinesForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "LinesForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'// This is free and unencumbered software released into the public domain.
'// For more information, please refer to  https://github.com/hongwenjun

'// ������� VBA_UserForm
Private Const TOOLNAME As String = "LYVBA"
Private Const SECTION As String = "LinesForm"

'// �û����ڳ�ʼ��
Private Sub UserForm_Initialize()

  With Me
    .StartUpPosition = 0
    .Left = Val(GetSetting(TOOLNAME, SECTION, "form_left", 900))
    .Top = Val(GetSetting(TOOLNAME, SECTION, "form_top", 200))
  End With

End Sub


'// �رմ���ʱ���洰��λ��
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    saveFormPos True
End Sub

'// ���洰��λ�úͼ��ش���λ��
Sub saveFormPos(bDoSave As Boolean)
  If bDoSave Then 'save position
    SaveSetting TOOLNAME, SECTION, "form_left", Me.Left
    SaveSetting TOOLNAME, SECTION, "form_top", Me.Top
  End If
End Sub

Private Sub MyPen_Click()
On Error GoTo ErrorHandler
  API.BeginOpt
  lines.Nodes_DrawLines
ErrorHandler:
  API.EndOpt
End Sub


'// ����Ҽ�Ctrl��������
Private Sub PenDrawLines_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
On Error GoTo ErrorHandler
  API.BeginOpt
  If Button = 2 Then
    lines.Draw_Multiple_Lines cdrAlignVCenter
    
  ElseIf Shift = fmCtrlMask Then
    lines.Draw_Multiple_Lines cdrAlignHCenter
  Else
    lines.Draw_Multiple_Lines 0
  End If
ErrorHandler:
  API.EndOpt
End Sub


'''////  ɵ�ϻ�����  ////'''
Private Sub TOP_ALIGN_BT_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
  If Button = 2 Then
    Tools.Simple_Train_Arrangement 3#
  ElseIf Shift = fmCtrlMask Then
    Tools.Simple_Train_Arrangement 0#
  Else
    Tools.Simple_Train_Arrangement Set_Space_Width
  End If
End Sub

'''////  ɵ�Ͻ�������  ////'''
Private Sub LEFT_ALIGN_BT_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
  If Button = 2 Then
    Tools.Simple_Ladder_Arrangement 3#
  ElseIf Shift = fmCtrlMask Then
    Tools.Simple_Ladder_Arrangement 0#
  Else
    Tools.Simple_Ladder_Arrangement Set_Space_Width
  End If
End Sub


Private Sub MakeBox_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
  On Error GoTo ErrorHandler
  API.BeginOpt
  
  Dim size As Variant
  size = input_box_lwh
  l = size(0): w = size(1): h = size(2): b = size(3)
  If b = 0 Then b = 15
  
  If Button = 2 Then
    box.Simple_box_five l, w, h, b
  ElseIf Shift = fmCtrlMask Then
    box.Simple_box_four l, w, h, b
  Else
    box.Simple_box_three l, w, h, b
  End If
  
ErrorHandler:
  API.EndOpt
End Sub

Private Sub Cmd_3D_Click()
  box.Simple_3Deffect
End Sub


'// �ǶȺ���ת����, �����ת���Ҽ���ת
Private Sub Rotate_Shapes_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
  If Button = 2 Then   '// �Ҽ��Ĵ���
    Shapes_Rotate -90
  ElseIf Shift = fmCtrlMask Then     '// ����Ĵ���
    Shapes_Rotate 90
  Else    '// CTRL�Ĵ���
    Shapes_Rotate -45
  End If
End Sub

'// �ƶ������ƣ������������������ƣ����ֻ�ƶ����Ҽ��Ƿ����򣬰�CTRL �Ǹ��Ƶ�
Private Sub Move_Left_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
  If Button = 2 Then   '// �Ҽ��Ĵ���
    move_shapes 100, 0
  ElseIf Shift = fmCtrlMask Then     '// ����Ĵ���
    move_shapes -100, 0
  Else    '// CTRL�Ĵ���
    Duplicate_shapes -100, 0
  End If
End Sub

Private Sub Move_Up_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
  If Button = 2 Then   '// �Ҽ��Ĵ���
    move_shapes 0, -100
  ElseIf Shift = fmCtrlMask Then     '// ����Ĵ���
    move_shapes 0, 100
  Else    '// CTRL�Ĵ���
    Duplicate_shapes 0, 100
  End If
End Sub

'// ������ߺ�ˮƽ��
Private Sub Ruler_Measuring_BT_Click()
  '// �Ƕ�תƽ
  Angle_to_Horizon
End Sub

'// ѡ������ƽ������
Private Sub Average_Distance_BT_Click()
  Average_Distance
End Sub

Private Sub chkAutoDistribute_Click()
  AutoDistribute_Key = chkAutoDistribute.Value
End Sub

'// ���񹤾�
Private Sub MirrorLine_Click()
  Mirror_ByGuide
End Sub

'// ƽ���߹��� CTRL �����þ���
Private Sub ParallelLines_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
  Dim sp As Double
  text = GlobalUserData("SpaceWidth", 1)
  sp = Val(text)
  If Button = 2 Then   '// �Ҽ��Ĵ���
    Create_Parallel_Lines -sp
  ElseIf Shift = fmCtrlMask Then     '// ����Ĵ���
    Create_Parallel_Lines sp
  Else    '// CTRL�Ĵ���
    Create_Parallel_Lines Set_Space_Width
  End If
End Sub

'// ��Ǿ���ο���
Private Sub Set_Guide_Click()
  Set_Guides_Name
End Sub
