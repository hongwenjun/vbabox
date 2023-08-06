# vbabox

![](https://github.com/hongwenjun/vbabox/blob/main/img/vbabox.webp)

CorelDRAW VBA 插件 简易的长宽高盒子插件和源码和步骤原理

https://www.bilibili.com/video/BV1MF411f7qu/

### 前两天在本论坛看到网友，问怎么绘制如下图简单的封套盒子。所以有了思路，就动手写了代码

![](https://lyvba.com/wp-content/uploads/2023/08/box1.png)

```bas
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
```

### 代码中的公式，不用初中知识应该就可以理解的。就是计算矩形多大，绘制4条线，把线移动到正确的位置

![](https://lyvba.com/wp-content/uploads/2023/08/box2.png)
```
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
```

### 绘制线条的代码写成函数调用，可以设置颜色或者粗细
```
'// 画一条线，设置轮廓色 M100
Private Function DrawLine(X1, Y1, X2, Y2) As Shape
  Set DrawLine = ActiveLayer.CreateLineSegment(X1, Y1, X2, Y2)
  DrawLine.Outline.SetProperties Color:=CreateCMYKColor(0, 100, 0, 0)
End Function
```
