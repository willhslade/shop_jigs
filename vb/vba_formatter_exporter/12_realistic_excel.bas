Option Explicit

Private Const SHEET_NAME As String = "Data"

Public Sub RefreshReport()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim total As Double
    Dim dict As Object

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    On Error GoTo CleanFail

    Set ws = ThisWorkbook.Worksheets(SHEET_NAME)
    Set dict = CreateObject("Scripting.Dictionary")
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    For r = 2 To lastRow
        If Len(ws.Cells(r, 1).Value) = 0 Then GoTo NextRow
        If Not dict.Exists(ws.Cells(r, 1).Value) Then
            dict.Add ws.Cells(r, 1).Value, 0#
        End If
        dict(ws.Cells(r, 1).Value) = dict(ws.Cells(r, 1).Value) + ws.Cells(r, 3).Value
        total = total + ws.Cells(r, 3).Value
NextRow:
    Next r

    With ws.Range("F1")
        .Value = "Total"
        .Font.Bold = True
        .Offset(1, 0).Value = total
        .Offset(1, 0).NumberFormat = "#,##0.00"
    End With

    Dim k As Variant
    For Each k In dict.Keys
        Debug.Print k & ": " & Format$(dict(k), "0.00")
    Next k

CleanExit:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Exit Sub

CleanFail:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Report"
    Resume CleanExit
End Sub

Public Function SafeDivide(ByVal numerator As Double, ByVal denominator As Double, _
        Optional ByVal fallback As Double = 0#) As Double
    If denominator = 0 Then
        SafeDivide = fallback
    Else
        SafeDivide = numerator / denominator
    End If
End Function

Public Function LastUsedRow(ByVal ws As Worksheet, Optional ByVal col As Long = 1) As Long
    LastUsedRow = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
End Function
