Option Explicit

Dim startingTime As Single

Sub startTimer()
    startingTime = Timer
End Sub

Function endTimer() As Single
    endTimer = Timer - startingTime
End Function
