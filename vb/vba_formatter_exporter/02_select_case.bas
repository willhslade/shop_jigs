Public Function Describe(ByVal v As Variant) As String
Select Case TypeName(v)
Case "Integer", "Long"
Describe = "whole"
Case "Double"
Describe = "float"
Case Else
Describe = "other"
End Select
End Function

Sub Nested(ByVal x As Long, ByVal y As Long)
    Select Case x
        Case Is < 0
            Select Case y
                Case 0
                    Debug.Print "neg zero"
                Case 1 To 5
                    Debug.Print "neg small"
                Case Else
                    Debug.Print "neg big"
            End Select
        Case 0
            Debug.Print "zero"
        Case Is >= 10, 7
            Debug.Print "big"
    End Select

    Select Case True
    Case x > 5 And y > 5
        Debug.Print "both"
    Case x > 5
        Debug.Print "x"
    End Select

    Select Case x: Case 1: Debug.Print "one": End Select
End Sub
