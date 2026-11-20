Sub Handler()
    On Error GoTo ErrHandler
    Dim x As Long
    x = 1 / 0
    On Error Resume Next
    x = 2
    On Error GoTo 0
    GoTo Cleanup

Retry:
    x = x + 1
    If x < 3 Then GoTo Retry

Cleanup:
    Exit Sub

ErrHandler:
    Select Case Err.Number
        Case 11
            Resume Next
        Case 13
            Resume Retry
        Case Else
            MsgBox Err.Description
    End Select
    Resume Cleanup
10  x = 1
20  If x = 1 Then GoTo 10
End Sub

Function Safe() As Long
    On Error GoTo Fail
    Safe = 1
Done: Exit Function
Fail:
    Safe = -1
    Resume Done
End Function
