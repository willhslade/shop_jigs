Option Explicit
Option Compare Text
Option Private Module

Private Const MAX_ITEMS As Long = 100
Public Const APP_NAME As String = "Demo"
Dim mCount As Long, mName As String
Private mItems() As Variant

Public Sub MessyIndent()
Dim i As Long, j As Long
   For i = 1 To 10
        If i Mod 2 = 0 Then
   For j = 1 To 3
      Debug.Print i, j
            Next j
      ElseIf i = 5 Then
  Debug.Print "five"
         Else
                Debug.Print "odd"
      End If
Next i

            Do While mCount < 5
      mCount = mCount + 1
   Loop

Do
    mCount = mCount - 1
Loop Until mCount <= 0

  While mCount < 3
mCount = mCount + 1
Wend

    With ActiveSheet
.Range("A1").Value = 1
        With .Range("B1")
    .Font.Bold = True
  End With
        End With
End Sub

Function Add(a As Long, b As Long) As Long
Add = a + b
End Function

Private Function Fact(ByVal n As Long) As Long
    If n <= 1 Then
        Fact = 1
    Else
        Fact = n * Fact(n - 1)
    End If
End Function
