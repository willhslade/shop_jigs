Sub InlineColons()
    Dim i As Long, x As Long
    For i = 1 To 3: x = x + i: Next i
    For i = 1 To 3: x = x + i
    Next i
    Do: x = x + 1: Loop While x < 5
    With Range("A1"): .Value = 1: End With
    x = 1: i = 2
    Dim s As String: s = "a:b"
    Debug.Print x: Debug.Print i
    Call Foo(a:=1, b:=2): x = 5
    If x = 1 Then x = 2: i = 3
    While x < 5: x = x + 1: Wend
End Sub

Sub OneLiner(): End Sub
Function OneLineFunc() As Long: OneLineFunc = 1: End Function
Private Sub Another(): Debug.Print "hi": End Sub

Sub AfterOneLiners()
    Debug.Print "still level one"
End Sub
