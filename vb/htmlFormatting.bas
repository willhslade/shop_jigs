Attribute VB_Name = "htmlFormatting"
Option Explicit


'http://www.ozgrid.com/News/excel-named-ranges.htm
Public Function GetAddress(HyperlinkCell As Range)
'If using Excel 97 use the WorksheetFunction.Substitute in place of Replace'    GetAddress = WorksheetFunction.Substitute
     GetAddress = Replace _
    (HyperlinkCell.Hyperlinks(1).Address, "mailto:", "")
End Function

Public Sub DeleteAllPictures()
'http://www.ozgrid.com/forum/showthread.php?t=63687
    ActiveSheet.Shapes.SelectAll
    Selection.Delete
End Sub

Public Sub DeleteAllHyperlinks()
'ME
    ActiveSheet.Hyperlinks.Delete
End Sub


