Attribute VB_Name = "modFormatterVBE"
Option Explicit

' ---------------------------------------------------------------------------
' VBE integration. Late bound (Object) so no VBIDE reference is needed.
' Requires: File > Options > Trust Center > Trust Center Settings >
'           Macro Settings > "Trust access to the VBA project object model".
' ---------------------------------------------------------------------------

Private Const MENU_TAG As String = "VBAFormatterMenu"
Private Const CT_ACTIVEX_DESIGNER As Long = 11
Private Const PP_LOCKED As Long = 1
Private Const MAX_DP_CELLS As Long = 4000000

' ===========================================================================
' Macros (menu targets)
' ===========================================================================

Public Sub FormatSelection()

    Dim Pane As Object
    Dim CM As Object
    Dim L1 As Long, C1 As Long, L2 As Long, C2 As Long
    Dim Opt As TFormatOptions

    On Error GoTo EH

    Set Pane = Application.VBE.ActiveCodePane
    If Pane Is Nothing Then
        MsgBox "Open a code window and select some lines first.", vbInformation, "VBA Formatter"
        Exit Sub
    End If

    Pane.GetSelection L1, C1, L2, C2
    If L1 = L2 And C1 = C2 Then
        FormatCurrentProcedure
        Exit Sub
    End If
    If L2 > L1 And C2 = 1 Then L2 = L2 - 1

    Set CM = Pane.CodeModule
    Opt = LoadOptions()
    FormatRange CM, L1, L2 - L1 + 1, Opt, True, True
    Application.StatusBar = "VBA Formatter: selection formatted."
    Exit Sub

EH:
    ReportError "Format selection"

End Sub

Public Sub FormatCurrentProcedure()

    Dim Pane As Object
    Dim CM As Object
    Dim L1 As Long, C1 As Long, L2 As Long, C2 As Long
    Dim Opt As TFormatOptions

    On Error GoTo EH

    Set Pane = Application.VBE.ActiveCodePane
    If Pane Is Nothing Then
        MsgBox "Open a code window and place the cursor in a procedure first.", vbInformation, "VBA Formatter"
        Exit Sub
    End If

    Pane.GetSelection L1, C1, L2, C2
    Set CM = Pane.CodeModule
    Opt = LoadOptions()
    FormatProcedureAt CM, L1, Opt
    Application.StatusBar = "VBA Formatter: procedure formatted."
    Exit Sub

EH:
    ReportError "Format procedure"

End Sub

Public Sub FormatCurrentModule()

    Dim Pane As Object
    Dim CM As Object
    Dim Comp As Object
    Dim Proj As Object
    Dim Opt As TFormatOptions
    Dim Backup As String
    Dim Changed As Boolean

    On Error GoTo EH

    Set Pane = Application.VBE.ActiveCodePane
    If Pane Is Nothing Then
        MsgBox "Open a code window first.", vbInformation, "VBA Formatter"
        Exit Sub
    End If

    Set CM = Pane.CodeModule
    Set Comp = CM.Parent
    Set Proj = Comp.Collection.Parent
    If Proj.Protection = PP_LOCKED Then
        MsgBox "This project is locked.", vbExclamation, "VBA Formatter"
        Exit Sub
    End If
    If CM.CountOfLines = 0 Then Exit Sub

    Opt = LoadOptions()
    If Opt.BackupBeforeBulk Then Backup = BackupProject(Proj, Comp.Name)

    Changed = FormatRange(CM, 1, CM.CountOfLines, Opt, False, False)

    If Changed Then
        Application.StatusBar = "VBA Formatter: " & Comp.Name & " formatted."
    Else
        Application.StatusBar = "VBA Formatter: " & Comp.Name & " already formatted."
    End If
    Exit Sub

EH:
    ReportError "Format module"

End Sub

Public Sub FormatActiveProject()

    Dim Proj As Object
    Dim Comp As Object
    Dim CM As Object
    Dim Opt As TFormatOptions
    Dim Backup As String
    Dim Total As Long
    Dim Changed As Long
    Dim Failed As String
    Dim Msg As String

    On Error GoTo EH

    Set Proj = Application.VBE.ActiveVBProject
    If Proj Is Nothing Then
        MsgBox "Select a project in the Project Explorer first.", vbInformation, "VBA Formatter"
        Exit Sub
    End If
    If Proj.Protection = PP_LOCKED Then
        MsgBox "This project is locked.", vbExclamation, "VBA Formatter"
        Exit Sub
    End If
    If IsOwnProject(Proj) Then
        If MsgBox("This is the formatter's own project. Format it anyway?", _
                  vbYesNo + vbQuestion, "VBA Formatter") <> vbYes Then Exit Sub
    End If
    If MsgBox("Format every module in '" & Proj.Name & "'?", _
              vbYesNo + vbQuestion, "VBA Formatter") <> vbYes Then Exit Sub

    Opt = LoadOptions()
    If Opt.BackupBeforeBulk Then Backup = BackupProject(Proj)

    For Each Comp In Proj.VBComponents
        If Comp.Type <> CT_ACTIVEX_DESIGNER Then
            Set CM = Comp.CodeModule
            If CM.CountOfLines > 0 Then
                Total = Total + 1
                On Error Resume Next
                If FormatRange(CM, 1, CM.CountOfLines, Opt, False, False) Then Changed = Changed + 1
                If Err.Number <> 0 Then
                    Failed = Failed & Comp.Name & " "
                    Err.Clear
                End If
                On Error GoTo EH
            End If
        End If
    Next Comp

    Msg = Changed & " of " & Total & " modules changed."
    If Len(Failed) > 0 Then Msg = Msg & vbCrLf & "Failed: " & Failed
    If Len(Backup) > 0 Then Msg = Msg & vbCrLf & vbCrLf & "Backup: " & Backup
    MsgBox Msg, vbInformation, "VBA Formatter"
    Exit Sub

EH:
    ReportError "Format project"

End Sub

Public Sub EditFormatterOptions()

    On Error GoTo EH
    OpenOptionsFile
    Exit Sub

EH:
    ReportError "Open options"

End Sub

' ===========================================================================
' Range handling
' ===========================================================================

Private Function FormatProcedureAt(ByVal CM As Object, ByVal CaretLine As Long, _
                                   ByRef Opt As TFormatOptions) As Boolean

    Dim Nm As String
    Dim PK As Long
    Dim First As Long
    Dim Cnt As Long

    On Error Resume Next
    Nm = CM.ProcOfLine(CaretLine, PK)
    If Err.Number <> 0 Then
        Nm = ""
        Err.Clear
    End If
    On Error GoTo 0

    If Len(Nm) > 0 Then
        First = CM.ProcStartLine(Nm, PK)
        Cnt = CM.ProcCountLines(Nm, PK)
    Else
        First = 1
        Cnt = CM.CountOfDeclarationLines
    End If
    If Cnt <= 0 Then Exit Function

    FormatProcedureAt = FormatRange(CM, First, Cnt, Opt, False, True)

End Function

' Formats CM lines FirstLine .. FirstLine+LineCount-1 and writes back only what changed.
' IsFragment: keep the leading/trailing blank lines of the range and never add Option Explicit.
' BaseFromIndent: take the starting indent level from the first non-blank line.
Private Function FormatRange(ByVal CM As Object, ByVal FirstLine As Long, ByVal LineCount As Long, _
                             ByRef Opt As TFormatOptions, ByVal BaseFromIndent As Boolean, _
                             ByVal IsFragment As Boolean) As Boolean

    Dim Norm As String
    Dim NewText As String
    Dim Lns() As String
    Dim N As Long
    Dim I As Long
    Dim Lead As Long
    Dim Trail As Long
    Dim Level As Long

    Norm = Replace(Replace(CM.Lines(FirstLine, LineCount), vbCrLf, vbLf), vbCr, vbLf)
    Lns = Split(Norm, vbLf)
    N = UBound(Lns) + 1
    If N > LineCount Then N = LineCount
    If N <= 0 Then Exit Function
    ReDim Preserve Lns(0 To N - 1)

    I = 0
    Do While I < N
        If IsBlankLine(Lns(I)) Then
            Lead = Lead + 1
            I = I + 1
        Else
            Exit Do
        End If
    Loop
    If Lead >= N Then Exit Function

    I = N - 1
    Do While I >= 0
        If IsBlankLine(Lns(I)) Then
            Trail = Trail + 1
            I = I - 1
        Else
            Exit Do
        End If
    Loop

    If BaseFromIndent Then Level = LeadingLevel(Lns, N, Opt)

    If IsFragment Then
        NewText = FormatWithOptions(Join(Lns, vbLf), Opt, Level, True)
        For I = 1 To Lead
            NewText = vbCrLf & NewText
        Next I
        For I = 1 To Trail
            NewText = NewText & vbCrLf
        Next I
    Else
        NewText = FormatWithOptions(Join(Lns, vbLf), Opt, -1, False)
    End If

    FormatRange = ApplyChanges(CM, FirstLine, Lns, NewText)

End Function

Private Function IsBlankLine(ByVal S As String) As Boolean

    IsBlankLine = (Len(Trim$(Replace(S, vbTab, " "))) = 0)

End Function

Private Function LeadingLevel(ByRef Lns() As String, ByVal N As Long, ByRef Opt As TFormatOptions) As Long

    Dim I As Long
    Dim J As Long
    Dim W As Long
    Dim Ch As String

    For I = 0 To N - 1
        If Not IsBlankLine(Lns(I)) Then
            For J = 1 To Len(Lns(I))
                Ch = Mid$(Lns(I), J, 1)
                If Ch = " " Then
                    W = W + 1
                ElseIf Ch = vbTab Then
                    W = W + Opt.IndentSize
                Else
                    Exit For
                End If
            Next J
            LeadingLevel = W \ Opt.IndentSize
            Exit Function
        End If
    Next I

End Function

' ===========================================================================
' Minimal write-back: common prefix/suffix trimmed, then an LCS diff of the middle,
' applied bottom-up with ReplaceLine/DeleteLines/InsertLines so unchanged lines keep
' their breakpoints and bookmarks.
' ===========================================================================

Private Function ApplyChanges(ByVal CM As Object, ByVal FirstLine As Long, _
                              ByRef Old() As String, ByVal NewText As String) As Boolean

    Dim Nw() As String
    Dim M As Long, K As Long, P As Long, S As Long
    Dim Mo As Long, Mn As Long
    Dim I As Long, J As Long, Q As Long
    Dim T() As Long
    Dim SegOld() As Long, SegDel() As Long, SegNew() As Long, SegIns() As Long
    Dim SegN As Long
    Dim GD As Long, GI As Long
    Dim IsMatch As Boolean

    Nw = Split(Replace(Replace(NewText, vbCrLf, vbLf), vbCr, vbLf), vbLf)
    M = UBound(Old) + 1
    K = UBound(Nw) + 1

    Do While P < M And P < K
        If Old(P) = Nw(P) Then P = P + 1 Else Exit Do
    Loop
    Do While S < M - P And S < K - P
        If Old(M - 1 - S) = Nw(K - 1 - S) Then S = S + 1 Else Exit Do
    Loop

    Mo = M - P - S
    Mn = K - P - S
    If Mo = 0 And Mn = 0 Then Exit Function
    ApplyChanges = True

    If Mo = 0 Or Mn = 0 Or CDbl(Mo + 1) * CDbl(Mn + 1) > MAX_DP_CELLS Then
        ApplySegment CM, FirstLine + P, Old, P, Mo, Nw, P, Mn
        Exit Function
    End If

    ReDim T(0 To Mo, 0 To Mn)
    For I = Mo - 1 To 0 Step -1
        For J = Mn - 1 To 0 Step -1
            If Old(P + I) = Nw(P + J) Then
                T(I, J) = T(I + 1, J + 1) + 1
            ElseIf T(I + 1, J) >= T(I, J + 1) Then
                T(I, J) = T(I + 1, J)
            Else
                T(I, J) = T(I, J + 1)
            End If
        Next J
    Next I

    ReDim SegOld(1 To Mo + Mn + 1)
    ReDim SegDel(1 To Mo + Mn + 1)
    ReDim SegNew(1 To Mo + Mn + 1)
    ReDim SegIns(1 To Mo + Mn + 1)

    I = 0
    J = 0
    Do While I < Mo Or J < Mn
        IsMatch = False
        If I < Mo And J < Mn Then IsMatch = (Old(P + I) = Nw(P + J))

        If IsMatch Then
            If GD + GI > 0 Then
                SegN = SegN + 1
                SegOld(SegN) = I - GD
                SegDel(SegN) = GD
                SegNew(SegN) = J - GI
                SegIns(SegN) = GI
                GD = 0
                GI = 0
            End If
            I = I + 1
            J = J + 1
        ElseIf J >= Mn Then
            GD = GD + 1
            I = I + 1
        ElseIf I >= Mo Then
            GI = GI + 1
            J = J + 1
        ElseIf T(I, J + 1) >= T(I + 1, J) Then
            GI = GI + 1
            J = J + 1
        Else
            GD = GD + 1
            I = I + 1
        End If
    Loop
    If GD + GI > 0 Then
        SegN = SegN + 1
        SegOld(SegN) = I - GD
        SegDel(SegN) = GD
        SegNew(SegN) = J - GI
        SegIns(SegN) = GI
    End If

    For Q = SegN To 1 Step -1
        ApplySegment CM, FirstLine + P + SegOld(Q), Old, P + SegOld(Q), SegDel(Q), _
                     Nw, P + SegNew(Q), SegIns(Q)
    Next Q

End Function

Private Sub ApplySegment(ByVal CM As Object, ByVal ModLine As Long, _
                         ByRef OldArr() As String, ByVal OldStart As Long, ByVal DelCount As Long, _
                         ByRef NewArr() As String, ByVal NewStart As Long, ByVal InsCount As Long)

    Dim R As Long
    Dim Q As Long
    Dim Chunk As String

    R = DelCount
    If InsCount < R Then R = InsCount

    For Q = 0 To R - 1
        If OldArr(OldStart + Q) <> NewArr(NewStart + Q) Then
            CM.ReplaceLine ModLine + Q, NewArr(NewStart + Q)
        End If
    Next Q

    If DelCount > R Then
        CM.DeleteLines ModLine + R, DelCount - R
    ElseIf InsCount > R Then
        For Q = R To InsCount - 1
            If Q > R Then Chunk = Chunk & vbCrLf
            Chunk = Chunk & NewArr(NewStart + Q)
        Next Q
        CM.InsertLines ModLine + R, Chunk
    End If

End Sub

' ===========================================================================
' Backup, safety, errors
' ===========================================================================

Public Function BackupProject(ByVal Proj As Object, Optional ByVal OnlyName As String = "") As String

    Dim Dest As String
    Dim Comp As Object
    Dim Ext As String

    On Error GoTo EH

    Dest = Environ$("TEMP") & "\VBAFormatterBackup\" & Format$(Now, "yyyymmdd_hhnnss")
    EnsureFolder Dest

    For Each Comp In Proj.VBComponents
        If Len(OnlyName) = 0 Or StrComp(Comp.Name, OnlyName, vbTextCompare) = 0 Then
            Select Case Comp.Type
                Case 1: Ext = ".bas"
                Case 3: Ext = ".frm"
                Case Else: Ext = ".cls"
            End Select
            Comp.Export Dest & "\" & Comp.Name & Ext
        End If
    Next Comp

    BackupProject = Dest
    Exit Function

EH:
    BackupProject = ""

End Function

Public Function IsOwnProject(ByVal Proj As Object) As Boolean

    On Error Resume Next
    IsOwnProject = (StrComp(Proj.FileName, ThisWorkbook.FullName, vbTextCompare) = 0)

End Function

Public Sub ReportError(ByVal Action As String)

    Dim N As Long
    Dim D As String

    N = Err.Number
    D = Err.Description

    If InStr(1, D, "not trusted", vbTextCompare) > 0 Or InStr(1, D, "programmatic access", vbTextCompare) > 0 Then
        MsgBox "Excel is blocking access to the VBA project." & vbCrLf & vbCrLf & _
               "Enable it under File > Options > Trust Center > Trust Center Settings > " & _
               "Macro Settings > 'Trust access to the VBA project object model'.", _
               vbExclamation, "VBA Formatter"
    Else
        MsgBox Action & " failed:" & vbCrLf & D & " (" & N & ")", vbExclamation, "VBA Formatter"
    End If

End Sub

' ===========================================================================
' VBE menu
' ===========================================================================

Public Sub InstallMenu()

    Dim Bar As Object
    Dim Pop As Object
    Dim Ctx As Object

    On Error GoTo EH

    RemoveMenu

    Set Bar = Application.VBE.CommandBars("Menu Bar")
    Set Pop = Bar.Controls.Add(10, , , , True)
    Pop.Caption = "VBA For&matter"
    Pop.Tag = MENU_TAG

    AddButton Pop, "Format &Selection", "FormatSelection"
    AddButton Pop, "Format &Procedure", "FormatCurrentProcedure"
    AddButton Pop, "Format &Module", "FormatCurrentModule"
    AddButton Pop, "Format &Project", "FormatActiveProject"
    AddButton Pop, "&Export Project Source...", "ExportProjectSource"
    AddButton Pop, "&Import Project Source...", "ImportProjectSource"
    AddButton Pop, "Check Project &Safety (no changes)", "CheckActiveProjectSafety"
    AddButton Pop, "Run &Corpus Tests...", "RunCorpusTests"
    AddButton Pop, "&Options...", "EditFormatterOptions"

    Set Ctx = Application.VBE.CommandBars("Code Window")
    AddButton Ctx, "Format Procedure", "FormatCurrentProcedure"
    Exit Sub

EH:
    Debug.Print "VBA Formatter menu not installed: " & Err.Description

End Sub

Public Sub RemoveMenu()

    On Error Resume Next
    RemoveTagged "Menu Bar"
    RemoveTagged "Code Window"

End Sub

Private Sub RemoveTagged(ByVal BarName As String)

    Dim Bar As Object
    Dim Ctl As Object
    Dim Guard As Long

    Set Bar = Application.VBE.CommandBars(BarName)
    Do
        Set Ctl = Nothing
        Set Ctl = Bar.FindControl(, , MENU_TAG, , True)
        If Ctl Is Nothing Then Exit Do
        Ctl.Delete
        Guard = Guard + 1
    Loop While Guard < 20

End Sub

Private Sub AddButton(ByVal Parent As Object, ByVal Caption As String, ByVal MacroName As String)

    Dim Btn As Object

    Set Btn = Parent.Controls.Add(1, , , , True)
    Btn.Caption = Caption
    Btn.OnAction = "'" & ThisWorkbook.Name & "'!" & MacroName
    Btn.Tag = MENU_TAG

End Sub
