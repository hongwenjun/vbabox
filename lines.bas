Attribute VB_Name = "lines"
'// This is free and unencumbered software released into the public domain.
'// For more information, please refer to  https://github.com/hongwenjun

Sub start()
  LinesForm.Show 0
End Sub

Public Function Nodes_DrawLines()
  Dim sr As ShapeRange, sr_tmp As New ShapeRange, sr_lines As New ShapeRange
  Dim s As Shape, sh As Shape
  Dim nr As NodeRange
  Set sr = ActiveSelectionRange
  If sr.Count = 0 Then Exit Function
  
  For Each sh In sr
    Set nr = sh.Curve.Selection
    If nr.Count > 0 Then
      For Each n In nr
        Set s = ActiveLayer.CreateEllipse2(n.PositionX, n.PositionY, 0.5, 0.5)
        sr_tmp.Add s
      Next n
    End If
  Next sh
  
  '// 没有选择节点的情况，使用物件中心划线
  If sr_tmp.Count < 2 And sr.Count > 1 Then
    Set Line = DrawLine(sr(1), sr(2))
    sr_lines.Add Line
  End If

#If VBA7 Then
    sr_tmp.Sort "@shape1.left < @shape2.left"
#Else
    Set sr_tmp = X4_Sort_ShapeRange(sr_tmp, stlx)
#End If

  '// 使用 Count 遍历 shaperange 这种情况方便点
  For i = 1 To sr_tmp.Count - 1
    Set Line = DrawLine(sr_tmp(i), sr_tmp(i + 1))
    sr_lines.Add Line
  Next
  
  sr_tmp.Delete
  sr_lines.CreateSelection
End Function

Public Function Draw_Multiple_Lines(hv As cdrAlignType)
  Dim sr As ShapeRange, sr_lines As New ShapeRange
  Set sr = ActiveSelectionRange
  
  If sr.Count < 2 Then Exit Function
  
#If VBA7 Then
  If hv = cdrAlignVCenter Then
    '// 从左到右排序
    sr.Sort "@shape1.left < @shape2.left"
  ElseIf hv = cdrAlignHCenter Then
    '// 从上到下排序
    sr.Sort "@shape1.top < @shape2.top"
  End If
#Else
  '// X4_Sort_ShapeRange for CorelDRAW X4
  If hv = cdrAlignVCenter Then
    Set sr = X4_Sort_ShapeRange(sr, stlx)
  ElseIf hv = cdrAlignHCenter Then
    Set sr = X4_Sort_ShapeRange(sr, stty)
  End If
 
#End If

  For i = 1 To sr.Count - 1 Step 2
    Set Line = DrawLine(sr(i), sr(i + 1))
    sr_lines.Add Line
  Next
 
  sr_lines.CreateSelection
End Function

Public Function FirtLineTool()
  Dim sr As ShapeRange
  Set sr = ActiveSelectionRange
  If sr.Count > 1 Then
    Set Line = DrawLine(sr(1), sr(2))
  End If
End Function

Private Function DrawLine(ByVal s1 As Shape, ByVal s2 As Shape) As Shape
'// 创建线段方法在图层上的指定位置创建由单个线段组成的曲线。
 Set DrawLine = ActiveLayer.CreateLineSegment(s1.CenterX, s1.CenterY, s2.CenterX, s2.CenterY)

End Function

Private Sub Test()
  ActiveDocument.Unit = cdrMillimeter
  Set Rect = ActiveLayer.CreateRectangle(0, 0, 30, 30)
  Set ell = ActiveLayer.CreateEllipse2(50, 50, 10, 10)
  Set Line = DrawLine(Rect, ell)
End Sub
