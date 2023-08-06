Attribute VB_Name = "box"
Public Function Simple_box_three()
  ActiveDocument.Unit = cdrMillimeter
  Dim sr As New ShapeRange, wing As New ShapeRange
  Dim sh As Shape
  l = 100: w = 50: h = 70: b = 15
  boxL = 2 * l + 2 * w + b: boxH = h
  l1x = w: l2x = w + l: l3x = 2 * w + l: l4x = 2 * (w + l)
  
  '// �����������¸Ǿ���
  Set mainRect = ActiveLayer.CreateRectangle(0, 0, boxL, boxH)
  Set topRect = ActiveLayer.CreateRectangle(0, 0, l, w)
  topRect.Move l1x, h
  Set bottomRect = ActiveLayer.CreateRectangle(0, 0, l, w)
  bottomRect.Move l3x, -w
  
  '// ����Box Բ�Ǿ��β��
  Set top_RoundRect = ActiveLayer.CreateRectangle(0, 0, l, b, 50, 50)
  top_RoundRect.Move l1x, h + w
  Set bottom_RoundRect = ActiveLayer.CreateRectangle(0, 0, l, b, 0, 0, 50, 50)
  bottom_RoundRect.Move l3x, -w - b
    
  '// ����box �ĸ����
  Set sh = DrawWing(ActiveLayer.CreateRectangle(0, 0, w, (w + b) / 2 - 2))
  wing.Add sh.Duplicate(0, h)
  wing.Add sh.Duplicate(l2x, h)
  wing.Add sh.Duplicate(0, -sh.SizeHeight)
  wing.Add sh.Duplicate(l2x, -sh.SizeHeight)
  wing(2).Flip cdrFlipHorizontal
  wing(3).Flip cdrFlipVertical
  wing(4).Rotate 180

  '// ��ӵ�����飬��������ɫ C100
  sr.Add mainRect: sr.Add topRect: sr.Add bottomRect
  sr.Add top_RoundRect: sr.Add bottom_RoundRect
  sr.AddRange wing: sh.Delete
  sr.SetOutlineProperties Color:=CreateCMYKColor(100, 0, 0, 0)
  
  '// ���Ƴߴ絶����
  Set sl1 = DrawLine(l1x, 0, l1x, h)
  Set sl2 = DrawLine(l2x, 0, l2x, h)
  Set sl3 = DrawLine(l3x, 0, l3x, h)
  Set sl4 = DrawLine(l4x, 0, l4x, h)
  
  '// ����box Ⱥ��
  sr.Add sl1: sr.Add sl2: sr.Add sl3: sr.Add sl4
  sr.CreateSelection: sr.Group
  
End Function

'// ��һ���ߣ���������ɫ M100
Private Function DrawLine(X1, Y1, X2, Y2) As Shape
  Set DrawLine = ActiveLayer.CreateLineSegment(X1, Y1, X2, Y2)
  DrawLine.Outline.SetProperties Color:=CreateCMYKColor(0, 100, 0, 0)
End Function


Private Function DrawWing(s As Shape) As Shape
    Dim sp As SubPath, crv As Curve
    Dim x As Double, y As Double
    x = s.SizeWidth: y = s.SizeHeight
    s.Delete
    
    '// ���� Box ��� Wing
    Set crv = Application.CreateCurve(ActiveDocument)
    Set sp = crv.CreateSubPath(0, 0)
    sp.AppendLineSegment 0, 4
    sp.AppendLineSegment 2, 6
    sp.AppendLineSegment 4, y - 2.5
    sp.AppendCurveSegment2 6.5, y, 4.1, y - 1.25, 5.1, y
    sp.AppendLineSegment x - 2, y
    sp.AppendLineSegment x - 2, 3
    sp.AppendLineSegment x, 0
    
    sp.Closed = True
    Set DrawWing = ActiveLayer.CreateCurve(crv)
End Function

Public Function Simple_box_one()
  ActiveDocument.Unit = cdrMillimeter
  l = 100: w = 50: h = 70: b = 15
  boxL = 2 * l + 2 * w + b
  boxH = h
  l1x = w
  l2x = w + l
  l3x = 2 * w + l
  l4x = 2 * (w + l)
  
  Set Rect = ActiveLayer.CreateRectangle(0, 0, boxL, boxH)
  Set sl1 = DrawLine(l1x, 0, l1x, h)
  Set sl2 = DrawLine(l2x, 0, l2x, h)
  Set sl3 = DrawLine(l3x, 0, l3x, h)
  Set sl4 = DrawLine(l4x, 0, l4x, h)
End Function

Public Function Simple_box_two()
  ActiveDocument.Unit = cdrMillimeter
  l = 100: w = 50: h = 70: b = 15
  boxL = 2 * l + 2 * w + b: boxH = h
  l1x = w: l2x = w + l: l3x = 2 * w + l: l4x = 2 * (w + l)
  
  Set mainRect = ActiveLayer.CreateRectangle(0, 0, boxL, boxH)
  
  Set topRect = ActiveLayer.CreateRectangle(0, 0, l, w)
  topRect.Move l1x, h
  Set bottomRect = ActiveLayer.CreateRectangle(0, 0, l, w)
  bottomRect.Move l3x, -w
  
  Set sl1 = DrawLine(l1x, 0, l1x, h)
  Set sl2 = DrawLine(l2x, 0, l2x, h)
  Set sl3 = DrawLine(l3x, 0, l3x, h)
  Set sl4 = DrawLine(l4x, 0, l4x, h)
End Function


Public Function Simple_3Deffect()
    Dim sr As ShapeRange    ' ���������Χ
    Set sr = ActiveSelectionRange   ' ѡ��3�����
  
    If sr.Count >= 3 Then
      ' // ����������������
      sr.Sort " @shape1.Top * 100 - @shape1.Left > @shape2.Top * 100 - @shape2.Left"
      
      sr(1).Stretch 0.951, 0.525      ' ����������������ͱ���
      sr(1).Skew 41.7, 7#
        
      sr(2).Stretch 0.951, 0.937      ' ����������������ͱ���
      sr(2).Skew 0#, 7#
      
      sr(3).Stretch 0.468, 0.937      ' ����������������ͱ���
      sr(3).Skew 0#, -45#
      
    End If
    
End Function
