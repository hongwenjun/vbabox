Attribute VB_Name = "AverageDistance"
Public AutoDistribute_Key As Boolean
Public first_StaticID As Long

'// 选择的物件平均距离
Public Function Average_Distance()
  On Error GoTo ErrorHandler
  API.BeginOpt
  
  Dim sr As ShapeRange
  Set sr = ActiveSelectionRange
  sr.Sort "@shape1.left<@shape2.left"

  Distribute_Shapes sr
  
ErrorHandler:
  API.EndOpt
End Function

Private Function Distribute_Shapes(sr As ShapeRange)
  Dim first As Double, last As Double
  Dim interval As Double, currentPoint As Double
  Dim total As Integer
  Dim sh As Shape
  
  first_StaticID = sr.FirstShape.StaticID
  total = sr.Count
  first = sr.FirstShape.CenterX
  last = sr.LastShape.CenterX
  interval = (last - first) / (total - 1)
  currentPoint = first


  For Each sh In sr
    sh.CenterY = sr.FirstShape.CenterY
    sh.CenterX = currentPoint
    currentPoint = currentPoint + interval
  Next sh
End Function
