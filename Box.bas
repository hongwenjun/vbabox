Attribute VB_Name = "box"
'// This is free and unencumbered software released into the public domain.
'// For more information, please refer to  https://github.com/hongwenjun

Public Function Simple_box_five(Optional ByVal l As Double, Optional ByVal w As Double, Optional ByVal h As Double, Optional ByVal b As Double = 15)
  Dim sr As New ShapeRange, wing As New ShapeRange, BottomWing As ShapeRange
  Dim sh As Shape
  l1x = w: l2x = w + l: l3x = 2 * w + l: l4x = 2 * (w + l)
  
  '// 绘制主体上下盖矩形
  Set mainRect_aw = ActiveLayer.CreateRectangle(0, 0, w, h)
  Set mainRect_al = ActiveLayer.CreateRectangle(0, 0, l, h)
  mainRect_al.Move l1x, 0
  Set mainRect_bw = ActiveLayer.CreateRectangle(0, 0, w, h)
  mainRect_bw.Move l2x, 0
  Set mainRect_bl = ActiveLayer.CreateRectangle(0, 0, l, h)
  mainRect_bl.Move l3x, 0

  Set topRect = ActiveLayer.CreateRectangle(0, 0, l, w)
  topRect.Move l1x, h

  '// 绘制Box 圆角矩形插口
  Set top_RoundRect = ActiveLayer.CreateRectangle(0, 0, l, b, 75, 75)
  top_RoundRect.Move l1x, h + w
  Set Bond = DrawBond(b, h, l4x, 0)

  Set SealLine = Draw_SealLine(l, l1x, h + w - 1)
  Set top_RoundRect = top_RoundRect.Weld(topRect, False, False)

  '// 绘制box 2个翅膀
  Set sh = DrawWing(w, (w + b) / 2 - 2)
  wing.Add sh.Duplicate(0, h)
  wing.Add sh.Duplicate(l2x, h)
  wing(2).Flip cdrFlipHorizontal

  '// 绘制 Box 底下翅膀 BottomWing
  Set BottomWing = DrawBottomWing(l, w, b)

  '// 添加到物件组，设置轮廓色 C100
  sr.Add mainRect_aw: sr.Add mainRect_al: sr.Add mainRect_bw: sr.Add mainRect_bl
  sr.Add Bond: sr.Add top_RoundRect
  sr.AddRange BottomWing
  sr.AddRange wing: sh.Delete
  sr.SetOutlineProperties Color:=CreateCMYKColor(100, 0, 0, 0)
  sr.Add SealLine
  
  sr.CreateSelection: sr.Group
End Function

Private Function DrawBottomWing(ByVal l As Double, ByVal w As Double, ByVal b As Double) As ShapeRange
  Dim sr As New ShapeRange, s As Shape
  Dim sp As SubPath, crv(3) As Curve
  
  w_2 = w / 2#
  block = w * 0.275 + w * ((l - w) / w * 0.15)
  bb = block + b: If bb < w_2 Then bb = w_2

  '// 绘制 Box 底下翅膀 BottomWing
  Set crv(1) = Application.CreateCurve(ActiveDocument)
  Set sp = crv(1).CreateSubPath(0, 0)
  sp.AppendLineSegment w_2, block
  sp.AppendLineSegment w_2, bb - 5
  sp.AppendCurveSegment2 w_2 + 5, bb, w_2, bb - 2.5, w_2 + 2.5, bb
  sp.AppendLineSegment w, bb
  sp.AppendLineSegment w, 0
  sp.Closed = True
  sr.Add ActiveLayer.CreateCurve(crv(1))
  
  Set crv(2) = Application.CreateCurve(ActiveDocument)
  Set sp = crv(2).CreateSubPath(0, 0)
  sp.AppendLineSegment w_2, block
  sp.AppendLineSegment w_2 + b - 5, block
  sp.AppendCurveSegment2 w_2 + b, block + 5, w_2 + b - 2.5, block, w_2 + b, block + 2.5
  sp.AppendLineSegment w_2 + b, l - block - 5
  sp.AppendCurveSegment2 w_2 + b - 5, l - block, w_2 + b, l - block - 2.5, w_2 + b - 2.5, l - block
  
  sp.AppendLineSegment w_2, l - block
  sp.AppendLineSegment 0, l
  sp.Closed = True
  sr.Add ActiveLayer.CreateCurve(crv(2))
  
  Set crv(3) = Application.CreateCurve(ActiveDocument)
  Set sp = crv(3).CreateSubPath(0, 0)
  sp.AppendLineSegment 0, l
  sp.AppendLineSegment w_2 + b, l
  sp.AppendLineSegment w_2 + b, l - block + 5
  sp.AppendCurveSegment2 w_2 + b - 5, l - block, w_2 + b, l - block + 2.5, w_2 + b - 2.5, l - block
  sp.AppendLineSegment w_2, l - block
  sp.AppendLineSegment w_2, block
  sp.AppendLineSegment w_2 + b - 5, block
  sp.AppendCurveSegment2 w_2 + b, block - 5, w_2 + b - 2.5, block, w_2 + b, block - 2.5
  sp.AppendLineSegment w_2 + b, 0
  sp.Closed = True
  sr.Add ActiveLayer.CreateCurve(crv(3))
  
  '// 移动到适合的地方
  sr(1).Move 0, -bb: sr(1).Rotate 180
  Set s = sr(1).Duplicate(0, 0): sr.Add s
  s.Flip cdrFlipHorizontal: s.Move w + l, 0
  
  sr(2).Rotate -90: sr(3).Rotate -90
  sr(2).LeftX = 2 * w + l: sr(3).LeftX = w
  sr(2).topY = 0: sr(3).topY = 0
  Set DrawBottomWing = sr
  
End Function


Public Function Simple_box_four(Optional ByVal l As Double, Optional ByVal w As Double, Optional ByVal h As Double, Optional ByVal b As Double = 15)
  Dim sr As New ShapeRange, wing As New ShapeRange
  Dim sh As Shape
  l1x = w: l2x = w + l: l3x = 2 * w + l: l4x = 2 * (w + l)
  
  '// 绘制主体上下盖矩形
  Set mainRect_aw = ActiveLayer.CreateRectangle(0, 0, w, h)
  Set mainRect_al = ActiveLayer.CreateRectangle(0, 0, l, h)
  mainRect_al.Move l1x, 0
  Set mainRect_bw = ActiveLayer.CreateRectangle(0, 0, w, h)
  mainRect_bw.Move l2x, 0
  Set mainRect_bl = ActiveLayer.CreateRectangle(0, 0, l, h)
  mainRect_bl.Move l3x, 0

  Set topRect = ActiveLayer.CreateRectangle(0, 0, l, w)
  topRect.Move l1x, h
  Set bottomRect = ActiveLayer.CreateRectangle(0, 0, l, w)
  bottomRect.Move l3x, -w
  
  '// 绘制Box 圆角矩形插口
  Set top_RoundRect = ActiveLayer.CreateRectangle(0, 0, l, b, 75, 75)
  top_RoundRect.Move l1x, h + w
  Set bottom_RoundRect = ActiveLayer.CreateRectangle(0, 0, l, b, 0, 0, 75, 75)
  bottom_RoundRect.Move l3x, -w - b
  Set Bond = DrawBond(b, h, l4x, 0)
  
  Set SealLine = Draw_SealLine(l, l1x, h + w - 1)
  Set SealLine2 = Draw_SealLine(l, l3x, -w - 1)
  SealLine2.Rotate 180
  
  '// 绘制box 四个翅膀
  Set sh = DrawWing(w, (w + b) / 2 - 2)
  wing.Add sh.Duplicate(0, h)
  wing.Add sh.Duplicate(l2x, h)
  wing.Add sh.Duplicate(0, -sh.SizeHeight)
  wing.Add sh.Duplicate(l2x, -sh.SizeHeight)
  wing(2).Flip cdrFlipHorizontal
  wing(3).Rotate 180
  wing(4).Flip cdrFlipVertical

  Set top_RoundRect = top_RoundRect.Weld(topRect, False, False)
  Set bottom_RoundRect = bottom_RoundRect.Weld(bottomRect, False, False)

  '// 添加到物件组，设置轮廓色 C100
  sr.Add mainRect_aw: sr.Add mainRect_al: sr.Add mainRect_bw: sr.Add mainRect_bl
  sr.Add Bond: sr.Add top_RoundRect: sr.Add bottom_RoundRect
  sr.AddRange wing: sh.Delete
  sr.SetOutlineProperties Color:=CreateCMYKColor(100, 0, 0, 0)
  sr.Add SealLine: sr.Add SealLine2
  sr.CreateSelection: sr.Group
  
End Function

Public Function input_box_lwh() As Variant
  Dim str, arr, n
  str = InputBox("请输入长x宽x高，使用空格 * x 间隔" & vbNewLine & "鼠标左键-右键-Ctrl三种样式", "盒子长宽高", "120 x 100 x 150 mm") & " "
  str = Newline_to_Space(str)

  ' 替换 mm x * 换行 TAB 为空格
  str = VBA.Replace(str, "mm", " ")
  str = VBA.Replace(str, "x", " ")
  str = VBA.Replace(str, "X", " ")
  str = VBA.Replace(str, "*", " ")

  '// 换行转空格 多个空格换成一个空格
  str = API.Newline_to_Space(str)

  arr = Split(str)
  arr(0) = Val(arr(0))
  arr(1) = Val(arr(1))
  arr(2) = Val(arr(2))
  arr(3) = Val(arr(3))
  input_box_lwh = arr
End Function

Public Function Simple_box_three(Optional ByVal l As Double, Optional ByVal w As Double, Optional ByVal h As Double, Optional ByVal b As Double = 15)
  ActiveDocument.Unit = cdrMillimeter
  Dim sr As New ShapeRange, wing As New ShapeRange
  Dim sh As Shape
  boxL = 2 * l + 2 * w + b: boxH = h
  l1x = w: l2x = w + l: l3x = 2 * w + l: l4x = 2 * (w + l)
  
  '// 绘制主体上下盖矩形
  Set mainRect = ActiveLayer.CreateRectangle(0, 0, boxL, boxH)
  Set topRect = ActiveLayer.CreateRectangle(0, 0, l, w)
  topRect.Move l1x, h
  Set bottomRect = ActiveLayer.CreateRectangle(0, 0, l, w)
  bottomRect.Move l3x, -w
  
  '// 绘制Box 圆角矩形插口
  Set top_RoundRect = ActiveLayer.CreateRectangle(0, 0, l, b, 75, 75)
  top_RoundRect.Move l1x, h + w
  Set bottom_RoundRect = ActiveLayer.CreateRectangle(0, 0, l, b, 0, 0, 75, 75)
  bottom_RoundRect.Move l3x, -w - b
  
  '// 绘制box 四个翅膀
  Set sh = DrawWing(w, (w + b) / 2 - 2)
  wing.Add sh.Duplicate(0, h)
  wing.Add sh.Duplicate(l2x, h)
  wing.Add sh.Duplicate(0, -sh.SizeHeight)
  wing.Add sh.Duplicate(l2x, -sh.SizeHeight)
  wing(2).Flip cdrFlipHorizontal
  wing(3).Rotate 180
  wing(4).Flip cdrFlipVertical

  Set top_RoundRect = top_RoundRect.Weld(topRect, False, False)
  Set bottom_RoundRect = bottom_RoundRect.Weld(bottomRect, False, False)
  
  '// 添加到物件组，设置轮廓色 C100
  sr.Add mainRect: sr.Add top_RoundRect: sr.Add bottom_RoundRect
  sr.AddRange wing: sh.Delete
  sr.SetOutlineProperties Color:=CreateCMYKColor(100, 0, 0, 0)
  
  '// 绘制尺寸刀痕线
  Set sl1 = DrawLine(l1x, 0, l1x, h)
  Set sl2 = DrawLine(l2x, 0, l2x, h)
  Set sl3 = DrawLine(l3x, 0, l3x, h)
  Set sl4 = DrawLine(l4x, 0, l4x, h)
  Set SealLine = Draw_SealLine(l, l1x, h + w - 1)
  Set SealLine2 = Draw_SealLine(l, l3x, -w - 1)
  SealLine2.Rotate 180
  
  '// 盒子box 群组
  sr.Add sl1: sr.Add sl2: sr.Add sl3: sr.Add sl4
  sr.Add SealLine: sr.Add SealLine2
  sr.CreateSelection: sr.Group
  
End Function

'// 画一条线，设置轮廓色 M100
Private Function DrawLine(x1, y1, x2, y2) As Shape
  Set DrawLine = ActiveLayer.CreateLineSegment(x1, y1, x2, y2)
  DrawLine.Outline.SetProperties Color:=CreateCMYKColor(0, 100, 0, 0)
End Function

'// 绘制封口线
Private Function Draw_SealLine(ByVal l As Double, ByVal move_x As Double, ByVal move_y As Double) As Shape
  Dim sp As SubPath, crv As Curve
  '// 绘制 Box 翅膀 Wing
  Set crv = Application.CreateCurve(ActiveDocument)
  Set sp = crv.CreateSubPath(0, 2)
  sp.AppendLineSegment 4, 2
  sp.AppendLineSegment 6, 0
  sp.AppendLineSegment l - 6, 0
  sp.AppendLineSegment l - 4, 2
  sp.AppendLineSegment l, 2
  sp.Closed = False
  Set Draw_SealLine = ActiveLayer.CreateCurve(crv)
  Draw_SealLine.Outline.SetProperties Color:=CreateCMYKColor(0, 100, 0, 0)
  Draw_SealLine.Move move_x, move_y
End Function

Private Function DrawWing(ByVal w As Double, ByVal h As Double) As Shape
    Dim sp As SubPath, crv As Curve
    Dim x As Double, y As Double
    x = w: y = h
    
    '// 绘制 Box 翅膀 Wing
    Set crv = Application.CreateCurve(ActiveDocument)
    Set sp = crv.CreateSubPath(0, 0)
    sp.AppendLineSegment 0, 4
    sp.AppendLineSegment 2, 6
    sp.AppendLineSegment 6, y - 2.5
    sp.AppendCurveSegment2 8.5, y, 6.2, y - 1.25, 7, y
    sp.AppendLineSegment x - 2, y
    sp.AppendLineSegment x - 2, 3
    sp.AppendLineSegment x, 0
    
    sp.Closed = True
    Set DrawWing = ActiveLayer.CreateCurve(crv)
End Function

Private Function DrawBond(ByVal w As Double, ByVal h As Double, ByVal move_x As Double, ByVal move_y As Double) As Shape
    Dim sp As SubPath, crv As Curve
    Dim x As Double, y As Double
    x = w: y = h
    
    '// 绘制 Box 粘合边 Bond
    Set crv = Application.CreateCurve(ActiveDocument)
    Set sp = crv.CreateSubPath(0, 0)
    sp.AppendLineSegment 0, y
    sp.AppendLineSegment x, y - 5
    sp.AppendLineSegment x, 5

    sp.Closed = True
    Set DrawBond = ActiveLayer.CreateCurve(crv)
    DrawBond.Move move_x, move_y
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
  Dim sr As ShapeRange            '// 定义物件范围
  Set sr = ActiveSelectionRange   '// 选择3个物件
  
  If sr.Count >= 3 Then
  
  '// 先上下再左右排序
#If VBA7 Then
    sr.Sort " @shape1.Top * 100 - @shape1.Left > @shape2.Top * 100 - @shape2.Left"
#Else
    Set ssr = X4_Sort_ShapeRange(sr, topWt_left)
#End If

    sr(1).Stretch 0.951, 0.525      ' 顶盖物件缩放修正和变形
    sr(1).Skew 41.7, 7#
      
    sr(2).Stretch 0.951, 0.937      ' 正面物件缩放修正和变形
    sr(2).Skew 0#, 7#
    
    sr(3).Stretch 0.468, 0.937      ' 侧面物件缩放修正和变形
    sr(3).Skew 0#, -45#
  End If
End Function
