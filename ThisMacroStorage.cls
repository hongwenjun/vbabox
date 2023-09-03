VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisMacroStorage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub GlobalMacroStorage_SelectionChange()
On Error GoTo ErrorHandler
  Dim n As Long
  Dim nr As NodeRange
  Dim sh As Shape
  
  If ActiveSelection.Shapes.Count > 0 Then
    n = 0
    For Each sh In ActiveSelection.Shapes
      If sh.Type = cdrCurveShape Then
        Set nr = sh.Curve.Selection
        n = n + nr.Count
      End If
    Next sh

    If n > 2 Then
        LinesForm.Caption = "Nodes: " & n
    ElseIf ActiveSelection.Shapes.Count > 1 Then
       LinesForm.Caption = "Select: " & ActiveSelection.Shapes.Count
      End If
  Else
      LinesForm.Caption = "LinesForm By Lanya"
  End If
  
  
  If ActiveSelection.Shapes.Count > 2 And AutoDistribute_Key Then
    Dim sr As ShapeRange
    Set sr = ActiveSelectionRange
    sr.Sort "@shape1.left<@shape2.left"
    If first_StaticID <> sr.FirstShape.StaticID Then
      Average_Distance
    End If
  End If
ErrorHandler:

End Sub
