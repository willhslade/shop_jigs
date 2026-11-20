Attribute VB_Name = "modFormatterSource"
Option Explicit

' ---------------------------------------------------------------------------
' Export / import of VBA source for version control. Late bound, works on the
' project that is active in the VBE (not on the add-in itself).
'
' Export : every component -> <folder>\Name.bas / .cls / .frm (+ .frx).
'          Sheets and ThisWorkbook are exported as .cls. Files are written only
'          when their content changed, so source control sees clean diffs.
' Import : .bas / .cls / .frm files from a folder.
'          Standard modules, classes and forms are replaced (old one is removed
'          only after the new one imported successfully).
'          Sheet / ThisWorkbook code cannot be re-imported as components, so
'          their code is swapped in place.
' Requires "Trust access to the VBA project object model".
' ---------------------------------------------------------------------------

Private Const CT_STDMODULE As Long = 1
Private Const CT_CLASS As Long = 2
Private Const CT_FORM As Long = 3
Private Const CT_DOCUMENT As Long = 100
Private Const FD_FOLDER_PICKER As Long = 4

' ===========================================================================
' Export
' ===========================================================================

Public Sub ExportProjectSource()

    Dim Proj As Object
    Dim Comp As Object
    Dim Dest As String
    Dim Tmp As String
    Dim Ext As String
    Dim F As String
    Dim Total As Long
    Dim Written As Long
    Dim Unchanged As Long
    Dim Failed As String

    On Error GoTo EH

    Set Proj = Application.VBE.ActiveVBProject
    If Proj Is Nothing Then
        MsgBox "Select a project in the Project Explorer first.", vbInformation, "VBA Formatter"
        Exit Sub
    End If

    Dest = DefaultSourceFolder(Proj)
    Dest = PickFolder("Export '" & Proj.Name & "' source to:", Dest)
    If Len(Dest) = 0 Then Exit Sub

    EnsureFolder Dest
    Tmp = Environ$("TEMP") & "\VBAFormatterExport_" & Format$(Now, "yyyymmdd_hhnnss")
    EnsureFolder Tmp

    For Each Comp In Proj.VBComponents
        Select Case Comp.Type
            Case CT_STDMODULE
                Ext = ".bas"
            Case CT_FORM
                Ext = ".frm"
            Case Else
                Ext = ".cls"
        End Select
        Total = Total + 1
        On Error Resume Next
        Comp.Export Tmp & "\" & Comp.Name & Ext
        If Err.Number <> 0 Then
            Failed = Failed & Comp.Name & " "
            Err.Clear
        End If
        On Error GoTo EH
    Next Comp

    F = Dir$(Tmp & "\*.*")
    Do While Len(F) > 0
        If CopyIfChanged(Tmp & "\" & F, Dest & "\" & F) Then
            Written = Written + 1
        Else
            Unchanged = Unchanged + 1
        End If
        F = Dir$()
    Loop

    CleanTemp Tmp

    ShowResult "Exported " & Total & " components to:" & vbCrLf & Dest & vbCrLf & vbCrLf & _
               Written & " file(s) written, " & Unchanged & " unchanged." & _
               IIf(Len(Failed) > 0, vbCrLf & "Failed: " & Failed, "")
    Exit Sub

EH:
    ReportError "Export source"

End Sub

' ===========================================================================
' Import
' ===========================================================================

Public Sub ImportProjectSource()

    Dim Proj As Object
    Dim Src As String
    Dim F As String
    Dim Files() As String
    Dim N As Long
    Dim I As Long
    Dim Status As String
    Dim Done As Long
    Dim Skipped As Long
    Dim Backup As String
    Dim Opt As TFormatOptions
    Dim Report As String

    On Error GoTo EH

    Set Proj = Application.VBE.ActiveVBProject
    If Proj Is Nothing Then
        MsgBox "Select a project in the Project Explorer first.", vbInformation, "VBA Formatter"
        Exit Sub
    End If
    If Proj.Protection = 1 Then
        MsgBox "This project is locked.", vbExclamation, "VBA Formatter"
        Exit Sub
    End If
    If IsOwnProject(Proj) Then
        MsgBox "That is the formatter's own project. Import into a different project.", vbExclamation, "VBA Formatter"
        Exit Sub
    End If

    Src = PickFolder("Import source into '" & Proj.Name & "' from:", DefaultSourceFolder(Proj))
    If Len(Src) = 0 Then Exit Sub

    ReDim Files(1 To 64)
    F = Dir$(Src & "\*.*")
    Do While Len(F) > 0
        Select Case LCase$(Right$(F, 4))
            Case ".bas", ".cls", ".frm"
                N = N + 1
                If N > UBound(Files) Then ReDim Preserve Files(1 To UBound(Files) * 2)
                Files(N) = F
        End Select
        F = Dir$()
    Loop

    If N = 0 Then
        MsgBox "No .bas, .cls or .frm files found in:" & vbCrLf & Src, vbInformation, "VBA Formatter"
        Exit Sub
    End If

    If MsgBox("Import " & N & " file(s) into '" & Proj.Name & "'?" & vbCrLf & _
              "Existing components with the same name will be replaced.", _
              vbYesNo + vbQuestion, "VBA Formatter") <> vbYes Then Exit Sub

    Opt = LoadOptions()
    If Opt.BackupBeforeBulk Then Backup = BackupProject(Proj)

    For I = 1 To N
        Status = ImportOne(Proj, Src & "\" & Files(I))
        If Left$(Status, 2) = "OK" Then
            Done = Done + 1
        Else
            Skipped = Skipped + 1
        End If
        Report = Report & vbCrLf & Files(I) & ": " & Status
    Next I

    ShowResult Done & " imported, " & Skipped & " skipped." & Report & _
               IIf(Len(Backup) > 0, vbCrLf & vbCrLf & "Backup: " & Backup, "")
    Exit Sub

EH:
    ReportError "Import source"

End Sub

' Returns "OK ..." or a reason for skipping.
Private Function ImportOne(ByVal Proj As Object, ByVal FilePath As String) As String

    Dim CompName As String
    Dim Existing As Object
    Dim Added As Object
    Dim OldName As String
    Dim Ext As String

    On Error GoTo EH

    CompName = ReadComponentName(FilePath)
    If Len(CompName) = 0 Then
        CompName = BaseName(FilePath)
    End If
    Ext = LCase$(Right$(FilePath, 4))

    Set Existing = Nothing
    On Error Resume Next
    Set Existing = Proj.VBComponents(CompName)
    On Error GoTo EH

    If Not Existing Is Nothing Then
        If Existing.Type = CT_DOCUMENT Then
            ReplaceDocumentCode Existing, FilePath
            ImportOne = "OK (code replaced)"
            Exit Function
        End If
    ElseIf LooksLikeDocument(FilePath) Then
        ImportOne = "skipped (sheet/workbook module not found in target)"
        Exit Function
    End If

    If Not Existing Is Nothing Then
        OldName = "zzOld_" & Format$(Now, "hhnnss") & "_" & CompName
        If Len(OldName) > 31 Then OldName = Left$(OldName, 31)
        Existing.Name = OldName
    End If

    On Error Resume Next
    Set Added = Proj.VBComponents.Import(FilePath)
    If Err.Number <> 0 Or Added Is Nothing Then
        Err.Clear
        If Not Existing Is Nothing Then Existing.Name = CompName
        ImportOne = "failed to import"
        Exit Function
    End If
    On Error GoTo EH

    If Not Existing Is Nothing Then Proj.VBComponents.Remove Existing

    ImportOne = "OK"
    Exit Function

EH:
    ImportOne = "failed: " & Err.Description

End Function

' Replaces the code of a Sheet / ThisWorkbook / document module with the code
' found in a .cls file (header and Attribute lines are dropped).
Private Sub ReplaceDocumentCode(ByVal Comp As Object, ByVal FilePath As String)

    Dim Lines() As String
    Dim N As Long
    Dim I As Long
    Dim Start As Long
    Dim Depth As Long
    Dim InHeader As Boolean
    Dim T As String
    Dim Code As String
    Dim CM As Object

    N = ReadLines(FilePath, Lines)
    If N = 0 Then Exit Sub

    I = 1
    Do While I <= N And Len(Trim$(Lines(I))) = 0
        I = I + 1
    Loop

    If I <= N Then
        If UCase$(Left$(Trim$(Lines(I)), 8)) = "VERSION " Then InHeader = True
    End If

    If InHeader Then
        I = I + 1
        Do While I <= N
            T = UCase$(Trim$(Lines(I)))
            If Left$(T, 6) = "BEGIN " Or T = "BEGIN" Then
                Depth = Depth + 1
            ElseIf T = "END" Then
                If Depth > 0 Then Depth = Depth - 1
                If Depth = 0 Then
                    I = I + 1
                    Exit Do
                End If
            End If
            I = I + 1
        Loop
    End If

    Do While I <= N
        If UCase$(Left$(Trim$(Lines(I)), 10)) = "ATTRIBUTE " Then
            I = I + 1
        Else
            Exit Do
        End If
    Loop
    Start = I

    Do While N >= Start And Len(Trim$(Lines(N))) = 0
        N = N - 1
    Loop

    For I = Start To N
        If I > Start Then Code = Code & vbCrLf
        Code = Code & Lines(I)
    Next I

    Set CM = Comp.CodeModule
    If CM.CountOfLines > 0 Then CM.DeleteLines 1, CM.CountOfLines
    If Len(Code) > 0 Then CM.AddFromString Code

End Sub

' ===========================================================================
' File helpers
' ===========================================================================

Private Function ReadComponentName(ByVal FilePath As String) As String

    Dim F As Integer
    Dim Ln As String
    Dim P As Long
    Dim Q As Long
    Dim Cnt As Long

    On Error GoTo Done
    F = FreeFile
    Open FilePath For Input As #F
    Do While Not EOF(F) And Cnt < 40
        Line Input #F, Ln
        Cnt = Cnt + 1
        If UCase$(Left$(Ln, 21)) = "ATTRIBUTE VB_NAME = """ Then
            P = InStr(Ln, """")
            Q = InStrRev(Ln, """")
            If Q > P Then ReadComponentName = Mid$(Ln, P + 1, Q - P - 1)
            Exit Do
        End If
    Loop

Done:
    On Error Resume Next
    If F <> 0 Then Close #F

End Function

Private Function LooksLikeDocument(ByVal FilePath As String) As Boolean

    Dim F As Integer
    Dim Ln As String
    Dim Cnt As Long

    On Error GoTo Done
    F = FreeFile
    Open FilePath For Input As #F
    Do While Not EOF(F) And Cnt < 40
        Line Input #F, Ln
        Cnt = Cnt + 1
        If InStr(1, Ln, "VB_Base", vbTextCompare) > 0 Then
            If InStr(Ln, "0{0002082") > 0 Then LooksLikeDocument = True
            Exit Do
        End If
    Loop

Done:
    On Error Resume Next
    If F <> 0 Then Close #F

End Function

Private Function ReadLines(ByVal FilePath As String, ByRef Lines() As String) As Long

    Dim F As Integer
    Dim Ln As String
    Dim N As Long

    ReDim Lines(1 To 256)
    F = FreeFile
    Open FilePath For Input As #F
    Do While Not EOF(F)
        Line Input #F, Ln
        N = N + 1
        If N > UBound(Lines) Then ReDim Preserve Lines(1 To UBound(Lines) * 2)
        Lines(N) = Ln
    Loop
    Close #F

    ReadLines = N

End Function

Private Function ReadBinary(ByVal FilePath As String) As String

    Dim F As Integer
    Dim Buf As String

    On Error GoTo Done
    F = FreeFile
    Open FilePath For Binary Access Read As #F
    If LOF(F) > 0 Then
        Buf = Space$(LOF(F))
        Get #F, , Buf
    End If
    ReadBinary = Buf

Done:
    On Error Resume Next
    If F <> 0 Then Close #F

End Function

' Copies Src over Dest only when Dest is missing or different. True if copied.
Private Function CopyIfChanged(ByVal Src As String, ByVal Dest As String) As Boolean

    If Len(Dir$(Dest)) > 0 Then
        If ReadBinary(Src) = ReadBinary(Dest) Then Exit Function
    End If

    FileCopy Src, Dest
    CopyIfChanged = True

End Function

Private Sub CleanTemp(ByVal Tmp As String)

    Dim F As String

    On Error Resume Next
    F = Dir$(Tmp & "\*.*")
    Do While Len(F) > 0
        Kill Tmp & "\" & F
        F = Dir$()
    Loop
    RmDir Tmp

End Sub

Private Function BaseName(ByVal FilePath As String) As String

    Dim S As String
    Dim P As Long

    S = Mid$(FilePath, InStrRev(FilePath, "\") + 1)
    P = InStrRev(S, ".")
    If P > 1 Then S = Left$(S, P - 1)
    BaseName = S

End Function

Private Function DefaultSourceFolder(ByVal Proj As Object) As String

    Dim Fn As String

    On Error Resume Next
    Fn = Proj.FileName
    On Error GoTo 0

    If Len(Fn) > 0 And InStrRev(Fn, "\") > 0 Then
        DefaultSourceFolder = Left$(Fn, InStrRev(Fn, "\") - 1) & "\src"
    Else
        DefaultSourceFolder = Environ$("USERPROFILE") & "\Documents\VBAFormatterExport\" & Proj.Name
    End If

End Function

' Folder picker; returns "" when cancelled. Falls back to the default folder if no dialog is available.
Private Function PickFolder(ByVal Title As String, ByVal StartFolder As String) As String

    Dim Dlg As Object
    Dim Start As String

    Start = StartFolder
    Do While Len(Start) > 3 And Len(Dir$(Start, vbDirectory)) = 0
        Start = Left$(Start, InStrRev(Start, "\") - 1)
    Loop

    On Error GoTo NoDialog
    Set Dlg = Application.FileDialog(FD_FOLDER_PICKER)
    Dlg.Title = Title
    If Len(Start) > 0 Then Dlg.InitialFileName = Start & "\"
    If Dlg.Show <> -1 Then Exit Function
    PickFolder = Dlg.SelectedItems(1)
    Exit Function

NoDialog:
    PickFolder = StartFolder

End Function

Private Sub ShowResult(ByVal Msg As String)

    MsgBox Msg, vbInformation, "VBA Formatter"

End Sub
