Attribute VB_Name = "sheetFunctions"
Option Explicit

Public Function sheetName(r As Range) As String
    sheetName = r.Parent.name
End Function

Public Function workBookName(r As Range) As String
    workBookName = r.Parent.Parent.name
End Function

Public Function sheetOrder(r As Range) As String
    sheetOrder = r.Parent.Index
End Function

