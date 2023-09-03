Attribute VB_Name = "RotateMoveDuplicate"
Public Function move_shapes(x As Double, y As Double)
  On Error GoTo ErrorHandler
  API.BeginOpt
  
  Dim sr As ShapeRange     '// ʹ�� ShapeRange ���Զ�����һ�����
  Set sr = ActiveSelectionRange   '// ѡ���������ʹ�� ActiveSelectionRange
  sr.Move x, y             '// Ĭ�ϵ�λ�� Ӣ�� �����ƶ�̫Զ��
  
ErrorHandler:
  API.EndOpt
End Function

Public Function Duplicate_shapes(x As Double, y As Double)
  On Error GoTo ErrorHandler
  API.BeginOpt
  
  Dim sr As ShapeRange
  Dim sr_copy As ShapeRange
  Set sr = ActiveSelectionRange
  Set sr_copy = sr.Duplicate(x, y)    '// Duplicate �����ƣ����ǰ���� = ��ֵ����Ҫ���� (x,y)
  sr_copy.CreateSelection

ErrorHandler:
  API.EndOpt
End Function

'// ������ת�Ƕ�
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
