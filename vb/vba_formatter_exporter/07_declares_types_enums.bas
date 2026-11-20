Option Explicit

#If VBA7 Then
Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
#Else
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

Public Enum Color
    Red = 1
    Green = 2
    Blue = 4
End Enum

Private Type Point
    X As Long
    Y As Long
End Type

Public Type Rectangle
    TopLeft As Point
    BottomRight As Point
    Name As String * 20
End Type

Public Event Changed(ByVal OldValue As Long, ByVal NewValue As Long)
Private WithEvents mApp As Application
Implements IFoo

Dim a As Long, b As String, c() As Double
Public d As New Collection
Private Const PI As Double = 3.14159
Static counter As Long

Public Sub UseThem()
    Dim p As Point
    Dim rc As Rectangle
    p.X = 1: p.Y = 2
    RaiseEvent Changed(1, 2)
End Sub
