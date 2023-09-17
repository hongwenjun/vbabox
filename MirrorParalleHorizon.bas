Attribute VB_Name = "MirrorParalleHorizon"
'// �����˵������,Ϊ(x1,y1)��(x2,y2) ��ô��Ƕ�a��tanֵ: tana=(y2-y1)/(x2-x1)
'// ���Լ���arctan(y2-y1)/(x2-x1), �õ���Ƕ�ֵa
'// VB����atn(), ����ֵ�ǻ��ȣ���Ҫ ���� PI /180
Private Function lineangle(x1, y1, x2, y2) As Double
    pi = 4 * VBA.Atn(1)    '// ����Բ����
    If x2 = x1 Then
      lineangle = 90: Exit Function
    End If
    lineangle = VBA.Atn((y2 - y1) / (x2 - x1)) / pi * 180
  End Function
  
  '// �Ƕ�תƽ
  Public Function Angle_to_Horizon()
    On Error GoTo ErrorHandler
    API.BeginOpt
    Set sr = ActiveSelectionRange
    Set nr = sr.LastShape.DisplayCurve.Nodes.All
  
    If nr.Count = 2 Then
      x1 = nr.FirstNode.PositionX: y1 = nr.FirstNode.PositionY
      x2 = nr.LastNode.PositionX: y2 = nr.LastNode.PositionY
      a = lineangle(x1, y1, x2, y2): sr.Rotate -a
      sr.LastShape.Delete   '// ɾ���ο���
    End If
ErrorHandler:
    API.EndOpt
  End Function

'// �Զ���ת�Ƕ�
Public Function Auto_Rotation_Angle()
  On Error GoTo ErrorHandler
  API.BeginOpt
  
'  ActiveDocument.ReferencePoint = cdrCenter
  Set sr = ActiveSelectionRange
  Set nr = sr.LastShape.DisplayCurve.Nodes.All

  If nr.Count = 2 Then
    x1 = nr.FirstNode.PositionX: y1 = nr.FirstNode.PositionY
    x2 = nr.LastNode.PositionX: y2 = nr.LastNode.PositionY
    a = lineangle(x1, y1, x2, y2): sr.Rotate 90 + a
    sr.LastShape.Delete   '// ɾ���ο���
  End If
ErrorHandler:
  API.EndOpt
End Function

'// ��������
Public Function Exchange_Object()
  Set sr = ActiveSelectionRange
  If sr.Count = 2 Then
    x = sr.LastShape.CenterX: y = sr.LastShape.CenterY
    sr.LastShape.CenterX = sr.FirstShape.CenterX: sr.LastShape.CenterY = sr.FirstShape.CenterY
    sr.FirstShape.CenterX = x: sr.FirstShape.CenterY = y
  End If
End Function

'// ��Ǿ���ο���
Public Function Set_Guides_Name()
  On Error GoTo ErrorHandler
  API.BeginOpt
  Dim sr As ShapeRange, s As Shape
  Set sr = ActiveSelectionRange
  
  For Each s In sr
    s.name = "MirrorGuides"
  Next s

'// ��л���ܾ���������͸����70%
  With ActiveSelection.Transparency
    .ApplyUniformTransparency 70
 '   .AppliedTo = cdrApplyToFillAndOutline
 '   .MergeMode = cdrMergeNormal
  End With
  
ErrorHandler:
  API.EndOpt
End Function

'// �ο��߾���
Public Function Mirror_ByGuide()
  On Error GoTo ErrorHandler
  API.BeginOpt
  Dim sr As ShapeRange, gds As ShapeRange
  Set sr = ActiveSelectionRange
  Set gds = sr.Shapes.FindShapes(Query:="@name ='MirrorGuides'")
  
  If gds.Count > 0 Then
 '//   sr.RemoveRange gds
    Set nr = gds(1).DisplayCurve.Nodes.All
  Else
    Set nr = sr.LastShape.DisplayCurve.Nodes.All
 '//   sr.Remove sr.Count
  End If
  
  If nr.Count >= 2 Then
    byshape = False
    x1 = nr.FirstNode.PositionX: y1 = nr.FirstNode.PositionY
    x2 = nr.LastNode.PositionX: y2 = nr.LastNode.PositionY
    a = lineangle(x1, y1, x2, y2)  '// �ο��ߺ�ˮƽ�ļн� a
    
    ang = 90 - a    '// �������ת�Ƕ�
   Set s = sr.Group
      With s
        Set s_copy = .Duplicate   '// �������������Ȼ�� x1,y1 �� ��ת
        
        .RotationCenterX = x1
        .RotationCenterY = y1
        .Rotate ang
        If Not byshape Then
            lx = .LeftX
            .Stretch -1#, 1#    '// ͨ��������ɾ���
            .LeftX = lx
            .Move (x1 - .LeftX) * 2 - .SizeWidth, 0
            .RotationCenterX = x1     '// ֮ǰ��Ϊ������ת���ĵ㷴�ˣ����û���
            .RotationCenterY = y1
            .Rotate -ang
        End If
        .RotationCenterX = .CenterX   '// ���û���ת���ĵ�Ϊ�������
        .RotationCenterY = .CenterY
        .Ungroup
        s_copy.Ungroup
      End With
  End If

ErrorHandler:
  API.EndOpt
End Function

'// �������ƽ����
Public Function Create_Parallel_Lines(space As Double)
  On Error GoTo ErrorHandler
  API.BeginOpt
  
  Dim sr As ShapeRange
  Set sr = ActiveSelectionRange
  sr.CreateParallelCurves 1, space

ErrorHandler:
  API.EndOpt
End Function
