VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisMacroStorage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Declare PtrSafe Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public sreg As New ShapeRange

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

  If ActiveSelection.Shapes.Count = 1 Then
    '// ¼ì²âCtrl:Alt:Shift¼ü×´Ì¬ 17-18-16
    
    If scankey() = 17 Then
      If sreg.Exists(ActiveShape) Then sreg.Remove sreg.IndexOf(ActiveShape)
      sreg.Add ActiveShape
      LinesForm.Caption = "ActiveShape add SREG! Count:" & sreg.Count
    End If
    
    If scankey() = 18 Then
      sreg.RemoveAll
      LinesForm.Caption = "SREG is Removed!"
    End If
    
    If scankey() = 16 Then
      sreg.CreateSelection
    End If
  
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

Private Function scankey() As Long
    Dim ctrlPressed As Boolean
    Dim shiftPressed As Boolean
    Dim altPressed As Boolean
    
    
    ' ¼ì²âCtrl¼üµÄ×´Ì¬  ' ¼ì²âShift¼üµÄ×´Ì¬   ' ¼ì²âAlt¼üµÄ×´Ì¬
    ctrlPressed = GetAsyncKeyState(17) And &H8000
    shiftPressed = GetAsyncKeyState(16) And &H8000
    altPressed = GetAsyncKeyState(18) And &H8000
    
    scankey = 0
    If altPressed Then scankey = 18
    If shiftPressed Then scankey = 16
    If ctrlPressed Then scankey = 17
End Function
