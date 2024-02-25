Attribute VB_Name = "userFunctions"
Option Explicit

'http://www.mrexcel.com/forum/excel-questions/85754-how-do-i-get-windows-login-name-visual-basic-applications.html
Public Function userName()
    userName = Environ("UserName")
End Function

Public Function userDomain()
    userDomain = Environ("UserDomain")
End Function
    
Public Function ComputerName()
    ComputerName = Environ("ComputerName")
End Function

