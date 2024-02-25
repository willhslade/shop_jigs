Attribute VB_Name = "colourMapWorksheet"
Option Explicit
'http://spreadsheetpage.com/index.php/site/tip/creating_a_worksheet_map/
Public Sub QuickMap()
    If TypeName(ActiveSheet) <> "Worksheet" Then Exit Sub

'   Create object variables for cell subsets
    On Error Resume Next
    Dim FormulaCells As Range
    Dim TextCells As Range
    Dim NumberCells  As Range
    Set FormulaCells = Range("A1").SpecialCells _
      (xlFormulas, xlNumbers + xlTextValues + xlLogical)
    Set TextCells = Range("A1").SpecialCells(xlConstants, xlTextValues)
    Set NumberCells = Range("A1").SpecialCells(xlConstants, xlNumbers)
    On Error GoTo 0

'   Add a new sheet and format it
    Sheets.Add
    With Cells
        .ColumnWidth = 2
        .Font.size = 8
        .HorizontalAlignment = xlCenter
    End With
    
    Application.ScreenUpdating = False

'   Do the formula cells
    Dim Area As Range
    If Not IsEmpty(FormulaCells) Then
        For Each Area In FormulaCells.Areas
            With ActiveSheet.Range(Area.Address)
                .Value = "F"
                .Interior.ColorIndex = 3
            End With
        Next Area
    End If
   
'   Do the text cells
    If Not IsEmpty(TextCells) Then
        For Each Area In TextCells.Areas
            With ActiveSheet.Range(Area.Address)
                .Value = "T"
                .Interior.ColorIndex = 4
            End With
        Next Area
    End If
    
'   Do the numeric cells
    If Not IsEmpty(NumberCells) Then
        For Each Area In NumberCells.Areas
            With ActiveSheet.Range(Area.Address)
                .Value = "N"
                .Interior.ColorIndex = 6
            End With
        Next Area
    End If
End Sub

