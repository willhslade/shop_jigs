Sub StringsAndComments()
    Dim s As String
    s = "If x Then"
    s = "End Sub"
    s = "' not a comment"
    s = "He said ""Next"" and ""End If"""
    s = "a=b,c<>d"
    s = "path\to\file.txt": s = "second"
    s = ""
    s = """"
    s = "tab	inside"
    Debug.Print "Select Case"   ' If this were code it would indent
    ' If this comment were code it would break the indent
    ' Function Foo()
    ' End Sub
    Rem End If
    Rem
    Dim r As Long ' trailing ' double ' comment
    r = 1 ' x = 1 : y = 2
    '    leading spaces inside the comment stay
    '
    ''' triple
    s = "a" '''
End Sub

' module level comment between procedures
' End Function

Sub AfterComments()
    Debug.Print "ok"
End Sub
