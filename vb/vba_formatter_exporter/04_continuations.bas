Sub Continuations()
    Dim s As String
    s = "a" & _
        "b" & _
        "c"
    s = Left(s, 1) & Mid(s, 2, _
        1)
    Debug.Print s; _
        s; s
    MsgBox "Hello", _
        vbInformation + vbOKOnly, _
        "Title"
    Call Foo(1, _
             2, _
             3)
    Foo a:=1, _
        b:=2
    If s = "a" And _
       s <> "b" And _
       Len(s) > 0 Then
        Debug.Print "multi"
    End If
    x = Array(1, 2, _
      3, 4)
    With Range("A1")
        .Value = 1 + _
            2
    End With
    For i = 1 To _
        10
        Debug.Print i
    Next i
    Select Case s & _
        "x"
        Case "ax"
            Debug.Print "ok"
    End Select
    Dim q As String: q = "ends with underscore _"
    q = "a _" & _
        "b"
End Sub

Private Sub Foo(Optional a As Long, Optional b As Long, Optional c As Long)
End Sub
