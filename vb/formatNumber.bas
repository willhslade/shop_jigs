Attribute VB_Name = "formatNumber"
Option Explicit

Public Sub switchFormatting()
Attribute switchFormatting.VB_ProcData.VB_Invoke_Func = "e\n14"
'ctrl-e
    If Selection.NumberFormat = "#,##0.00,," Then
        Selection.NumberFormat = "#,##0"
    Else
        Selection.NumberFormat = "#,##0.00,,"
    End If
End Sub

