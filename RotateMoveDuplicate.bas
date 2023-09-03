Attribute VB_Name = "RotateMoveDuplicate"
Public Function move_shapes(x As Double, y As Double)
  On Error GoTo ErrorHandler
  API.BeginOpt
  
  Dim sr As ShapeRange     '// 使用 ShapeRange 可以多个物件一起操作
  Set sr = ActiveSelectionRange   '// 选择物件队列使用 ActiveSelectionRange
  sr.Move x, y             '// 默认单位是 英寸 所以移动太远了
  
ErrorHandler:
  API.EndOpt
End Function

Public Function Duplicate_shapes(x As Double, y As Double)
  On Error GoTo ErrorHandler
  API.BeginOpt
  
  Dim sr As ShapeRange
  Dim sr_copy As ShapeRange
  Set sr = ActiveSelectionRange
  Set sr_copy = sr.Duplicate(x, y)    '// Duplicate 是再制，如果前面有 = 赋值，就要加上 (x,y)
  sr_copy.CreateSelection

ErrorHandler:
  API.EndOpt
End Function

'// 批量旋转角度
Public Function Shapes_Rotate(angle As Double)
  On Error GoTo ErrorHandler
  API.BeginOpt
  
  ActiveDocument.ReferencePoint = cdrCenter
  Dim sr As ShapeRange
  Set sr = ActiveSelectionRange
  For Each s In sr
    s.Rotate angle
  Next
  
ErrorHandler:
  API.EndOpt
End Function
