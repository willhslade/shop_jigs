Attribute VB_Name = "pasteValuesNewSheet"
Option Explicit
 
Private Sub PasteValueSheets()
'paste values all sheets
'UNTESTED
'http://www.teachexcel.com/excel-help/excel-how-to.php?i=178025
    Dim NewName As String
    Dim ws As Worksheet
     
    With Application
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
        
        On Error GoTo ErrCatcher
    
    'loop over all visible sheets and paste values
        For Each ws In ActiveWorkbook.Worksheets
            If ws.Visible Then
                Application.CutCopyMode = False
                ws.Activate
                ws.Cells.Copy
                ws.[A1].PasteSpecial Paste:=xlValues
                ws.Cells.Hyperlinks.Delete
                Cells(1, 1).Select
            End If
        Next ws
        Application.CutCopyMode = False
        Cells(1, 1).Select
        
    'delete all hidden sheets
        For Each ws In ActiveWorkbook.Worksheets
            If Not ws.Visible Then
                ws.Delete
            End If
        Next ws
        
        NewName = ActiveWorkbook.name & Format(Date, "yyyy-mm-dd") & "_psv.xlsx"
        
        ActiveWorkbook.SaveCopyAs ActiveWorkbook.path & "\" & NewName
        'ActiveWorkbook.Close SaveChanges:=False
         
        .Calculation = xlCalculationAutomatic
        .ScreenUpdating = True
    End With
Exit Sub

ErrCatcher:
End Sub
