Attribute VB_Name = "lastRowLastColumn"
Option Explicit
'http://spreadsheetpage.com/index.php/site/tip/determining_the_last_non_empty_cell_in_a_column_or_row/

Public Function LASTINCOLUMN(rngInput As Range)
    Dim WorkRange As Range
    Dim i As Long, CellCount As Long
    Application.Volatile
    Set WorkRange = rngInput.Columns(1).EntireColumn
    Set WorkRange = Intersect(WorkRange.Parent.UsedRange, WorkRange)
    CellCount = WorkRange.Count
    For i = CellCount To 1 Step -1
        If Not IsEmpty(WorkRange(i)) Then
            LASTINCOLUMN = WorkRange(i).Value
            Exit Function
        End If
    Next i
End Function

Public Function LASTINROW(rngInput As Range) As Variant
    Dim WorkRange As Range
    Dim i As Long, CellCount As Long
    Application.Volatile
    Set WorkRange = rngInput.Rows(1).EntireRow
    Set WorkRange = Intersect(WorkRange.Parent.UsedRange, WorkRange)
    CellCount = WorkRange.Count
    For i = CellCount To 1 Step -1
        If Not IsEmpty(WorkRange(i)) Then
            LASTINROW = WorkRange(i).Value
            Exit Function
        End If
    Next i
End Function

