Attribute VB_Name = "removeVBA"
Option Explicit

Public Sub clearImmediateWindow()
    Dim i As Integer
    For i = 1 To 10000
        Debug.Print ""
    Next i
End Sub
 
'http://www.vbaexpress.com/kb/getarticle.php?kb_id=93
Private Sub DeleteAllCode()
     
     'Trust Access To Visual Basics Project must be enabled.
     'From Excel: Tools | Macro | Security | Trusted Sources
     
    Dim x               As Integer
    Dim Proceed         As VbMsgBoxResult
    Dim Prompt          As String
    Dim Title           As String
     
    Prompt = "Are you certain that you want to delete all the VBA Code from " & _
    ActiveWorkbook.name & "?"
    Title = "Verify Procedure"
     
    Proceed = MsgBox(Prompt, vbYesNo + vbQuestion, Title)
    If Proceed = vbNo Then
        MsgBox "Procedure Canceled", vbInformation, "Procedure Aborted"
        Exit Sub
    End If
     
    On Error Resume Next
    With ActiveWorkbook.VBProject
        For x = .VBComponents.Count To 1 Step -1
            .VBComponents.Remove .VBComponents(x)
        Next x
        For x = .VBComponents.Count To 1 Step -1
            .VBComponents(x).CodeModule.DeleteLines _
            1, .VBComponents(x).CodeModule.CountOfLines
        Next x
    End With
    On Error GoTo 0
     
End Sub


