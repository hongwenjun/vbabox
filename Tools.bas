Attribute VB_Name = "Tools"
'// This is free and unencumbered software released into the public domain.
'// For more information, please refer to  https://github.com/hongwenjun

'// ���׻�����
Public Function Simple_Train_Arrangement(Space_Width As Double)
  API.BeginOpt
  Dim ssr As ShapeRange, s As Shape
  Dim cnt As Integer
  Set ssr = ActiveSelectionRange
  cnt = 1

#If VBA7 Then
'  ssr.sort " @shape1.top>@shape2.top"
  ssr.Sort " @shape1.left<@shape2.left"
#Else
' X4 ��֧�� ShapeRange.sort  ʹ�� lyvba32.dll �㷨������   2023.07.08
  Set ssr = X4_Sort_ShapeRange(ssr, stlx)
#End If

  ActiveDocument.ReferencePoint = cdrTopLeft
  For Each s In ssr
    '// �׶��� If cnt > 1 Then s.SetPosition ssr(cnt - 1).RightX, ssr(cnt - 1).BottomY
    '// �ĳɶ����� 2022-08-10
    ActiveDocument.ReferencePoint = cdrTopLeft + cdrBottomTop
    If cnt > 1 Then s.SetPosition ssr(cnt - 1).RightX + Space_Width, ssr(cnt - 1).TopY
    cnt = cnt + 1
  Next s

  API.EndOpt
End Function

'// ���׽�������
Public Function Simple_Ladder_Arrangement(Space_Width As Double)
  API.BeginOpt
  Dim ssr As ShapeRange, s As Shape
  Dim cnt As Integer
  Set ssr = ActiveSelectionRange
  cnt = 1

#If VBA7 Then
  ssr.Sort " @shape1.top>@shape2.top"
#Else
' X4 ��֧�� ShapeRange.sort  ʹ�� lyvba32.dll �㷨������   2023.07.08
  Set ssr = X4_Sort_ShapeRange(ssr, stty).ReverseRange
#End If


  ActiveDocument.ReferencePoint = cdrTopLeft
  For Each s In ssr
    If cnt > 1 Then s.SetPosition ssr(cnt - 1).LeftX, ssr(cnt - 1).BottomY - Space_Width
    cnt = cnt + 1
  Next s

  API.EndOpt
End Function
