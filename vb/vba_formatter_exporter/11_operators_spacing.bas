Sub Operators()
    Dim a As Long, b As Long, s As String, d As Date, v As Variant
    a=1:b=2
    a = a+b*2-1/3
    a = -a
    a = - 1
    b = a^2
    s = "a"&"b"
    If a<>b Then a=b
    If a<=b And a>=b Then a=1
    If a<b Or a>b Then a=1
    If a=b Then a=1
    d = #1/2/2020#
    d = #1/2/2020 3:04:05 PM#
    a = &HFF
    a = &O17
    a = 5 \ 2
    a = 5 Mod 2
    v = Array(1,2,3)
    v = Array( 1 , 2 , 3 )
    Foo 1,2,3
    Foo 1, , 3
    Foo a:=1,b:=2
    Call Foo(1,2 ,3)
    rs!Field = 1
    rs.Fields("x") = a
    s = Format(a, "0.00")
    a = (b + (a * 2)) / 3
    Debug.Print a;b;
    Debug.Print a, b
    Print #1, a; b
    Open "f.txt" For Input As #1
    Line Input #1, s
    Close #1
    If Not a = b Then a = 0
    x = a = b
    Mid(s, 1, 1) = "x"
    Set o = Nothing
    Let a = 3
    a = a And b Or Not b Xor a Eqv b Imp a
    s = s & vbCrLf & "x"
    a = a + 1 ' <>=<> in comment
End Sub

Sub Foo(Optional a As Long, Optional b As Long, Optional c As Long)
End Sub
