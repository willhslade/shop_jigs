Sub IfForms(ByVal a As Long, ByVal b As Long)
    If a = 1 Then b = 2
    If a = 1 Then b = 2 Else b = 3
    If a = 1 Then Exit Sub
    If a = 1 Then GoTo Done
    If a = 1 Then Debug.Print "x": Debug.Print "y"
    If a = 1 Then ' trailing comment makes this a block
        b = 4
    End If
    If a = 1 And _
       b = 2 Then
        b = 5
    ElseIf a = 2 Or _
           b = 3 Then
        b = 6
    ElseIf a = 3 Then: b = 7
    Else
        b = 8
    End If
    If a = 1 Then
        If b = 2 Then
            Debug.Print "nested"
        End If
    End If
    If IsNumeric(a) Then Debug.Print "IsNumeric then in name": Debug.Print "x"
    If Not (a = 1 Or b = 2) Then Debug.Print "not"
    If a = 1 Then If b = 2 Then Debug.Print "nested one-line"
    If a Then Else Debug.Print "else only"
Done:
    Debug.Print "done"
End Sub
