Option Explicit

Public Sub Blocks(ByVal n As Long)
Dim i As Long, total As Long
Dim s As String
For i = 1 To n
If i Mod 2 = 0 Then
total = total + i
ElseIf i Mod 3 = 0 Then
total = total - i
Else
total = total + 1
End If
Next i
Do While total > 0
total = total - 10
If total < 5 Then Exit Do
Loop
Do
total = total + 1
Loop Until total > 20
While total < 100
total = total * 2
Wend
With Worksheets(1)
.Range("A1").Value = total
With .Range("B1")
.Font.Bold = True
End With
End With
    For i = 1 To 3
        For j = 1 To 3
                If i = j Then
        Debug.Print i
                End If
        Next j
    Next i
End Sub

Private Function Twice(x As Long) As Long
Twice = x * 2
End Function
