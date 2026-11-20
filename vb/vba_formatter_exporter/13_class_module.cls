Option Explicit
Implements IShape

Private mWidth As Double
Private mHeight As Double
Public Event Resized()

Private Sub Class_Initialize()
mWidth = 1
End Sub

Private Sub Class_Terminate()
End Sub

Public Property Get Width() As Double
Width = mWidth
End Property

Public Property Let Width(ByVal v As Double)
mWidth = v
RaiseEvent Resized
End Property

Private Function IShape_Area() As Double
IShape_Area = mWidth * mHeight
End Function
