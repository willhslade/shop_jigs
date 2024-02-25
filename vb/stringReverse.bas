Attribute VB_Name = "stringReverse"
Option Explicit
    
Public Function ReverseString(Text As String)
'http://www.vbaexpress.com/kb/getarticle.php?kb_id=188
     
    ReverseString = StrReverse(Text)
     
End Function

Public Function REVERSE97(ByVal sCellContents As String) As String
'http://www.bettersolutions.com/excel/EIK284/LR521811611.htm
Dim ichar As Integer
   If Application.WorksheetFunction.IsNonText(sCellContents) = True Then
      REVERSE97 = VBA.CVErr(XlCVError.xlErrNA)
   Else
      For ichar = Len(sCellContents) To 1 Step -1
         REVERSE97 = REVERSE97 & Mid(sCellContents, ichar, 1)
      Next ichar
   End If
End Function
