Attribute VB_Name = "cellFunctions"
Option Explicit

Public Function CellType(c)
'http://spreadsheetpage.com/index.php/site/tip/determining_the_data_type_of_a_cell/
'   Returns the cell type of the upper left
'   cell in a range
    Application.Volatile
    Set c = c.Range("A1")
    Select Case True
        Case IsEmpty(c): CellType = "Blank"
        Case Application.IsText(c): CellType = "Text"
        Case Application.IsLogical(c): CellType = "Logical"
        Case Application.IsErr(c): CellType = "Error"
        Case IsDate(c): CellType = "Date"
        Case InStr(1, c.Text, ":") <> 0: CellType = "Time"
        Case IsNumeric(c): CellType = "Value"
    End Select
End Function

Function IsBold(rCell As Range)
    IsBold = rCell.Font.Bold
End Function

'see also rgb
Public Function cellColor(rCell As Range)
    'cellColor = rCell.Interior.ColorIndex
    cellColor = rCell.Interior.Color
End Function

Public Function fontColor(rCell As Range)
    'fontColor = rCell.Font.ColorIndex
    fontColor = rCell.Font.Color
End Function





