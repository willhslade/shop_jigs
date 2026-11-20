Option Explicit
Option Compare Text
Option Base 1
Option Private Module

Public Const MAX_ITEMS As Long = 100
Private Const NAME As String = "x"
Public gCount As Long
Private mItems() As Variant
Dim a As Long, b As String, c As Double
Public WithEvents btn As MSForms.CommandButton
Public Event Changed(ByVal oldValue As Long, ByVal newValue As Long)
Implements IFoo
Global gFlag As Boolean

Public Enum Color
    Red = 1
    Green = 2 ' green
    Blue = 4
End Enum

Private Type Point
x As Long
y As Long
tags(1 To 3) As String
End Type

Type Rect
TopLeft As Point
BottomRight As Point
End Type

Sub UsesStatic()
Static counter As Long
Dim arr(1 To 10) As Long
Dim m(1 To 3, 1 To 3) As Double
ReDim Preserve mItems(1 To 20)
counter = counter + 1
End Sub
