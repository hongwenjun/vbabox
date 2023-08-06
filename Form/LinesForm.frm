VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} LinesForm 
   Caption         =   "LinesForm"
   ClientHeight    =   855
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4725
   OleObjectBlob   =   "LinesForm.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "LinesForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MyPen_Click()
  lines.Nodes_DrawLines
End Sub


'// 左键右键Ctrl三键控制
Private Sub PenDrawLines_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
  If Button = 2 Then
    lines.Draw_Multiple_Lines cdrAlignVCenter
    
  ElseIf Shift = fmCtrlMask Then
    lines.Draw_Multiple_Lines cdrAlignHCenter
  Else
    lines.Draw_Multiple_Lines 0
  End If
End Sub



'''////  傻瓜火车排列  ////'''
Private Sub TOP_ALIGN_BT_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
  If Button = 2 Then
    Tools.Simple_Train_Arrangement 3#
  ElseIf Shift = fmCtrlMask Then
    Tools.Simple_Train_Arrangement 0#
  Else
    Tools.Simple_Train_Arrangement Set_Space_Width
  End If
End Sub

'''////  傻瓜阶梯排列  ////'''
Private Sub LEFT_ALIGN_BT_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
  If Button = 2 Then
    Tools.Simple_Ladder_Arrangement 3#
  ElseIf Shift = fmCtrlMask Then
    Tools.Simple_Ladder_Arrangement 0#
  Else
    Tools.Simple_Ladder_Arrangement Set_Space_Width
  End If
End Sub


Private Sub MakeBox_Click()
  box.Simple_box_three
End Sub

Private Sub Cmd_3D_Click()
  box.Simple_3Deffect
End Sub
