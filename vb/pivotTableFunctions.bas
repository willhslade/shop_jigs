Attribute VB_Name = "pivotTableFunctions"
Option Explicit
Sub DeleteMissingItems2002All() 'http://www.contextures.com/xlPivot04.html'prevents unused items in non-OLAP PivotTables
'pivot table tutorial by contextures.com
Dim pt As PivotTable
Dim ws As Worksheet
Dim pc As PivotCache
'change the settings
For Each ws In ActiveWorkbook.Worksheets
  For Each pt In ws.PivotTables
    pt.PivotCache.MissingItemsLimit = xlMissingItemsNone
  Next pt
Next ws

'refresh all the pivot caches
For Each pc In ActiveWorkbook.PivotCaches
  On Error Resume Next
  pc.Refresh
Next pc

End Sub
Public Sub DeleteOldItemsWB() 'XL 97/ XL 2000
'pivot table tutorial by contextures.com
'gets rid of unused items in PivotTable
' based on MSKB (202232)
Dim ws As Worksheet
Dim pt As PivotTable
Dim pf As PivotField
Dim pi As PivotItem

On Error Resume Next
For Each ws In ActiveWorkbook.Worksheets
  For Each pt In ws.PivotTables
    pt.RefreshTable
    pt.ManualUpdate = True
    For Each pf In pt.VisibleFields
      If pf.name <> "Data" Then
        For Each pi In pf.PivotItems
          If pi.RecordCount = 0 And _
            Not pi.IsCalculated Then
            pi.Delete
          End If
        Next pi
      End If
    Next pf
    pt.ManualUpdate = False    'pt.RefreshTable 'optional - might hang Excel
                 'if 2 or more pivot tables on one sheet
  Next pt
Next ws

End Sub

Public Sub RefreshAllPivotTables()
'http://stackoverflow.com/a/71084
'ALTERNATIVELY
'ThisWorkbook.RefreshAll
Dim pt As PivotTable
Dim ws As Worksheet

    For Each ws In ThisWorkbook.Worksheets
        For Each pt In ws.PivotTables
          pt.RefreshTable
        Next pt
    Next ws
End Sub

'http://www.ozgrid.com/VBA/pivot-table-refresh.htm
Public Sub RefreshAllWorksheetPivots()
Dim pt As PivotTable
    For Each pt In ActiveSheet.PivotTables
        pt.RefreshTable
    Next pt
End Sub

Public Sub debugPrintFieldNames()
    Call clearImmediateWindow
    Dim pivotName As String
    pivotName = ActiveCell.PivotCell.Parent.name
    
    Dim pTable As PivotTable
    Set pTable = ActiveCell.PivotTable
    
    Dim pvtField As PivotField
    For Each pvtField In pTable.PivotFields
        Debug.Print pvtField.name
    Next pvtField

    Set pvtField = Nothing
    Set pTable = Nothing
End Sub
