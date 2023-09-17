Attribute VB_Name = "MirrorParalleHorizon"
'// 两个端点的坐标,为(x1,y1)和(x2,y2) 那么其角度a的tan值: tana=(y2-y1)/(x2-x1)
'// 所以计算arctan(y2-y1)/(x2-x1), 得到其角度值a
'// VB中用atn(), 返回值是弧度，需要 乘以 PI /180
Private Function lineangle(x1, y1, x2, y2) As Double
    pi = 4 * VBA.Atn(1)    '// 计算圆周率
    If x2 = x1 Then
      lineangle = 90: Exit Function
    End If
    lineangle = VBA.Atn((y2 - y1) / (x2 - x1)) / pi * 180
  End Function
  
  '// 角度转平
  Public Function Angle_to_Horizon()
    On Error GoTo ErrorHandler
    API.BeginOpt
    Set sr = ActiveSelectionRange
    Set nr = sr.LastShape.DisplayCurve.Nodes.All
  
    If nr.Count = 2 Then
      x1 = nr.FirstNode.PositionX: y1 = nr.FirstNode.PositionY
      x2 = nr.LastNode.PositionX: y2 = nr.LastNode.PositionY
      a = lineangle(x1, y1, x2, y2): sr.Rotate -a
      sr.LastShape.Delete   '// 删除参考线
    End If
ErrorHandler:
    API.EndOpt
  End Function

'// 自动旋转角度
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
    sr.LastShape.Delete   '// 删除参考线
  End If
ErrorHandler:
  API.EndOpt
End Function

'// 交换对象
Public Function Exchange_Object()
  Set sr = ActiveSelectionRange
  If sr.Count = 2 Then
    x = sr.LastShape.CenterX: y = sr.LastShape.CenterY
    sr.LastShape.CenterX = sr.FirstShape.CenterX: sr.LastShape.CenterY = sr.FirstShape.CenterY
    sr.FirstShape.CenterX = x: sr.FirstShape.CenterY = y
  End If
End Function

'// 标记镜像参考线
Public Function Set_Guides_Name()
  On Error GoTo ErrorHandler
  API.BeginOpt
  Dim sr As ShapeRange, s As Shape
  Set sr = ActiveSelectionRange
  
  For Each s In sr
    s.name = "MirrorGuides"
  Next s

'// 感谢李总捐赠，定置透明度70%
  With ActiveSelection.Transparency
    .ApplyUniformTransparency 70
 '   .AppliedTo = cdrApplyToFillAndOutline
 '   .MergeMode = cdrMergeNormal
  End With
  
ErrorHandler:
  API.EndOpt
End Function

'// 参考线镜像
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
    a = lineangle(x1, y1, x2, y2)  '// 参考线和水平的夹角 a
    
    ang = 90 - a    '// 镜像的旋转角度
   Set s = sr.Group
      With s
        Set s_copy = .Duplicate   '// 复制物件保留，然后按 x1,y1 点 旋转
        
        .RotationCenterX = x1
        .RotationCenterY = y1
        .Rotate ang
        If Not byshape Then
            lx = .LeftX
            .Stretch -1#, 1#    '// 通过拉伸完成镜像
            .LeftX = lx
            .Move (x1 - .LeftX) * 2 - .SizeWidth, 0
            .RotationCenterX = x1     '// 之前因为镜像，旋转中心点反了，重置回来
            .RotationCenterY = y1
            .Rotate -ang
        End If
        .RotationCenterX = .CenterX   '// 重置回旋转中心点为物件中心
        .RotationCenterY = .CenterY
        .Ungroup
        s_copy.Ungroup
      End With
  End If

ErrorHandler:
  API.EndOpt
End Function

'// 物件建立平行线
Public Function Create_Parallel_Lines(space As Double)
  On Error GoTo ErrorHandler
  API.BeginOpt
  
  Dim sr As ShapeRange
  Set sr = ActiveSelectionRange
  sr.CreateParallelCurves 1, space

ErrorHandler:
  API.EndOpt
End Function
