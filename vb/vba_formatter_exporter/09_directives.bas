#Const DEBUGMODE = 1
#If VBA7 Then
Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long
#Else
Private Declare Function GetTickCount Lib "kernel32" () As Long
#End If

Sub Conditional()
#If DEBUGMODE = 1 Then
Debug.Print "debug"
#ElseIf Win64 Then
Debug.Print "64"
#Else
Debug.Print "release"
#End If
If x Then
#If VBA7 Then
y = 1
#End If
End If
End Sub
