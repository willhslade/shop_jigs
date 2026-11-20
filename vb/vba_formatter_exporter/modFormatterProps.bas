Attribute VB_Name = "modFormatterProps"
Option Explicit

' ---------------------------------------------------------------------------
' Safety tests for the formatter. None of them needs a hand-written expected output.
' For every piece of code and every option set they check that formatting
'
'   1. does not change the code      same tokens, ignoring only whitespace and line breaks
'   2. is stable                     formatting the result again changes nothing
'   3. is canonical                  re-indenting, tabs, trailing spaces or CR / CRLF line
'                                    endings in the INPUT never change the output
'   4. keeps VBA's limits            <= 24 continuations per statement, <= 1023 chars per line,
'                                    no trailing whitespace, indents are whole levels
'   5. indents correctly             every statement is at the level an independent
'                                    regex-based oracle expects (this is what catches drift)
'   6. keeps procedures balanced     Sub / Function / Property / Type / Enum start and end at
'                                    column 0 and never nest
'
' Entry points (run from the Macros dialog or the Immediate window):
'   RunCorpusTests          every .bas / .cls / .frm / .txt in a folder (default: tests\corpus)
'   CheckActiveProjectSafety  checks YOUR project in memory; writes nothing to it
'   RunPerformanceTest      times the formatter on a large module
'
' Needs Windows (late-bound VBScript.RegExp and Scripting.Dictionary).
' ---------------------------------------------------------------------------

Private mCache As Object        ' Scripting.Dictionary of compiled regular expressions

Private Const SET_COUNT As Long = 7

' ===========================================================================
' Entry points
' ===========================================================================

Public Sub RunCorpusTests()

    Dim Folder As String
    Dim F As String
    Dim Names() As String
    Dim N As Long
    Dim I As Long
    Dim Code As String
    Dim Report As String
    Dim Checks As Long
    Dim Fails As Long

    On Error GoTo EH

    Folder = ThisWorkbook.Path & "\tests\corpus"
    Folder = PickFolder("Folder with VBA test files (.bas .cls .frm .txt)", Folder)
    If Len(Folder) = 0 Then Exit Sub

    ReDim Names(1 To 64)
    F = Dir$(Folder & "\*.*")
    Do While Len(F) > 0
        Select Case LCase$(Right$(F, 4))
            Case ".bas", ".cls", ".frm", ".txt"
                N = N + 1
                If N > UBound(Names) Then ReDim Preserve Names(1 To UBound(Names) * 2)
                Names(N) = F
        End Select
        F = Dir$()
    Loop

    If N = 0 Then
        MsgBox "No test files found in:" & vbCrLf & Folder, vbInformation, "Formatter tests"
        Exit Sub
    End If

    Application.StatusBar = "Formatter tests running..."
    For I = 1 To N
        Code = ExtractCode(ReadFile(Folder & "\" & Names(I)))
        CheckCode Names(I), Code, True, Report, Checks, Fails
    Next I
    Application.StatusBar = False

    Finish "Corpus tests (" & N & " files)", Report, Checks, Fails
    Exit Sub

EH:
    Application.StatusBar = False
    MsgBox "Test run failed: " & Err.Description & " (" & Err.Number & ")", vbExclamation, "Formatter tests"

End Sub

Public Sub CheckActiveProjectSafety()

    Dim Proj As Object
    Dim Comp As Object
    Dim CM As Object
    Dim Report As String
    Dim Checks As Long
    Dim Fails As Long
    Dim Total As Long

    On Error GoTo EH

    Set Proj = Application.VBE.ActiveVBProject
    If Proj Is Nothing Then
        MsgBox "Select a project in the Project Explorer first.", vbInformation, "Formatter tests"
        Exit Sub
    End If

    Application.StatusBar = "Checking " & Proj.Name & "..."
    For Each Comp In Proj.VBComponents
        If Comp.Type <> 11 Then
            Set CM = Comp.CodeModule
            If CM.CountOfLines > 0 Then
                Total = Total + 1
                CheckCode Comp.Name, CM.Lines(1, CM.CountOfLines), False, Report, Checks, Fails
            End If
        End If
    Next Comp
    Application.StatusBar = False

    Finish "Project '" & Proj.Name & "' (" & Total & " modules, nothing was modified)", Report, Checks, Fails
    Exit Sub

EH:
    Application.StatusBar = False
    MsgBox "Check failed: " & Err.Description & " (" & Err.Number & ")", vbExclamation, "Formatter tests"

End Sub

Public Sub RunPerformanceTest()

    Dim Opt As TFormatOptions
    Dim Block As String
    Dim Big As String
    Dim I As Long
    Dim T0 As Double
    Dim Out As String
    Dim Lines As Long

    Opt = DefaultOptions()

    Block = "Public Function Sample" & "(ByVal a As Long, ByVal b As String) As Long" & vbCrLf & _
            "Dim i As Long" & vbCrLf & _
            "For i = 1 To 10" & vbCrLf & _
            "If i Mod 2 = 0 Then" & vbCrLf & _
            "Debug.Print ""even"", i ' comment" & vbCrLf & _
            "Else" & vbCrLf & _
            "Select Case i" & vbCrLf & _
            "Case 1, 3" & vbCrLf & _
            "a = a + i" & vbCrLf & _
            "Case Else" & vbCrLf & _
            "b = b & ""x""" & vbCrLf & _
            "End Select" & vbCrLf & _
            "End If" & vbCrLf & _
            "Next i" & vbCrLf & _
            "Sample = a" & vbCrLf & _
            "End Function" & vbCrLf & vbCrLf

    Dim Parts() As String
    ReDim Parts(1 To 1500)
    For I = 1 To 1500
        Parts(I) = Block
    Next I
    Big = Join(Parts, "")
    Lines = UBound(Split(Big, vbCrLf)) + 1

    T0 = Timer
    Out = FormatWithOptions(Big, Opt)
    T0 = Timer - T0

    MsgBox "Formatted " & Lines & " lines in " & Format$(T0, "0.00") & " s", vbInformation, "Formatter performance"
    Debug.Print "Performance: " & Lines & " lines, " & Format$(T0, "0.00") & " s, output " & Len(Out) & " chars"

End Sub

' ===========================================================================
' The checks
' ===========================================================================

Private Sub CheckCode(ByVal Label As String, ByVal Code As String, ByVal WithMutations As Boolean, _
                      ByRef Report As String, ByRef Checks As Long, ByRef Fails As Long)

    Dim S As Long
    Dim O As TFormatOptions
    Dim OName As String
    Dim Lbl As String
    Dim Out1 As String
    Dim Out2 As String
    Dim Probs As String
    Dim M As Long
    Dim MOut As String
    Dim KeyIn As String

    KeyIn = TokenKey(Code)

    For S = 1 To SET_COUNT
        O = OptionSet(S, OName)
        Lbl = Label & " [" & OName & "]"

        On Error Resume Next
        Out1 = FormatWithOptions(Code, O)
        If Err.Number <> 0 Then
            AddFail Report, Fails, Lbl & ": formatter raised error " & Err.Number & " " & Err.Description
            Err.Clear
            On Error GoTo 0
            GoTo NextSet
        End If
        Out2 = FormatWithOptions(Out1, O)
        If Err.Number <> 0 Then
            AddFail Report, Fails, Lbl & ": second pass raised error " & Err.Number & " " & Err.Description
            Err.Clear
            On Error GoTo 0
            GoTo NextSet
        End If
        On Error GoTo 0

        ' 2. stable
        Checks = Checks + 1
        If Out1 <> Out2 Then AddFail Report, Fails, Lbl & ": NOT STABLE - " & FirstDiff(Out1, Out2)

        ' 1. code unchanged
        Checks = Checks + 1
        If KeyIn <> TokenKey(Out1) Then AddFail Report, Fails, Lbl & ": CODE CHANGED - " & FirstTokenDiff(KeyIn, TokenKey(Out1))

        If S <> 5 Then
            ' 4 + 6. limits and procedure structure
            Checks = Checks + 1
            Probs = StructuralProblems(Out1, O)
            If Len(Probs) > 0 Then AddFail Report, Fails, Lbl & ": STRUCTURE - " & Probs

            ' 5. indentation oracle
            Checks = Checks + 1
            Probs = IndentProblems(Out1, O)
            If Len(Probs) > 0 Then AddFail Report, Fails, Lbl & ": INDENT - " & Probs
        End If

        ' 3. canonical
        If WithMutations And (S = 1 Or S = 2 Or S = 4) Then
            For M = 1 To 6
                Checks = Checks + 1
                On Error Resume Next
                MOut = FormatWithOptions(Mutate(Code, M), O)
                If Err.Number <> 0 Then
                    AddFail Report, Fails, Lbl & " " & MutationName(M) & ": formatter raised error " & Err.Description
                    Err.Clear
                    MOut = Out1
                End If
                On Error GoTo 0
                If MOut <> Out1 Then AddFail Report, Fails, Lbl & " input variant '" & MutationName(M) & "' gives different output - " & FirstDiff(Out1, MOut)
            Next M
        End If

NextSet:
    Next S

End Sub

Private Function OptionSet(ByVal Idx As Long, ByRef Nm As String) As TFormatOptions

    Dim O As TFormatOptions

    O = DefaultOptions()
    O.AddOptionExplicit = False

    Select Case Idx
        Case 1
            Nm = "default"
        Case 2
            Nm = "tabs"
            O.UseTabs = True
        Case 3
            Nm = "flat Case, indent 2"
            O.IndentCase = False
            O.IndentSize = 2
            O.ContinuationLevels = 2
        Case 4
            Nm = "always wrap, narrow"
            O.SignatureWrap = 2
            O.MaxLineLength = 50
        Case 5
            Nm = "raw (no spacing, no layout)"
            O.SignatureWrap = 0
            O.SpaceOperators = False
            O.SpaceCommas = False
            O.MaxBlankLines = 0
            O.BlankLinesBetweenProcs = -1
            O.BlankAfterDeclarations = False
        Case 6
            Nm = "aligned"
            O.AlignAs = True
            O.AlignComments = True
            O.CommentGap = 3
            O.MaxLineLength = 0
        Case 7
            Nm = "very narrow"
            O.SignatureWrap = 1
            O.MaxLineLength = 40
    End Select

    OptionSet = O

End Function

' ---------------------------------------------------------------------------
' 1. Tokens: strings, comments, words (case sensitive) and single punctuation characters.
'    Whitespace, blank lines and " _" continuations are ignored.
' ---------------------------------------------------------------------------

Private Function TokenKey(ByVal Code As String) As String

    Dim Lines() As String
    Dim Toks() As String
    Dim TN As Long
    Dim LastNL As Boolean
    Dim I As Long
    Dim J As Long
    Dim K As Long
    Dim N As Long
    Dim Ln As String
    Dim Ch As String

    Lines = JoinContinuations(Code)
    ReDim Toks(1 To 256)
    LastNL = True

    For I = LBound(Lines) To UBound(Lines)
        Ln = Lines(I)
        N = Len(Ln)
        J = 1
        Do While J <= N
            Ch = Mid$(Ln, J, 1)
            If Ch = " " Or Ch = vbTab Then
                J = J + 1
            ElseIf Ch = """" Then
                K = J + 1
                Do While K <= N
                    If Mid$(Ln, K, 1) = """" Then
                        If Mid$(Ln, K + 1, 1) = """" Then
                            K = K + 2
                        Else
                            Exit Do
                        End If
                    Else
                        K = K + 1
                    End If
                Loop
                AddTok Toks, TN, Mid$(Ln, J, K - J + 1)
                LastNL = False
                J = K + 1
            ElseIf Ch = "'" Then
                AddTok Toks, TN, RTrimWs(Mid$(Ln, J))
                LastNL = False
                J = N + 1
            ElseIf IsWordCh(Ch) Then
                K = J
                Do While K <= N
                    If IsWordCh(Mid$(Ln, K, 1)) Then K = K + 1 Else Exit Do
                Loop
                AddTok Toks, TN, Mid$(Ln, J, K - J)
                LastNL = False
                J = K
            Else
                AddTok Toks, TN, Ch
                LastNL = False
                J = J + 1
            End If
        Loop
        If Not LastNL Then
            AddTok Toks, TN, vbLf
            LastNL = True
        End If
    Next I

    If TN = 0 Then Exit Function
    ReDim Preserve Toks(1 To TN)
    TokenKey = Join(Toks, Chr$(1))

End Function

Private Sub AddTok(ByRef Toks() As String, ByRef TN As Long, ByVal T As String)

    TN = TN + 1
    If TN > UBound(Toks) Then ReDim Preserve Toks(1 To UBound(Toks) * 2)
    Toks(TN) = T

End Sub

' Joins " _" continuation lines (textual rule, applied to input and output alike).
Private Function JoinContinuations(ByVal Code As String) As String()

    Dim Src() As String
    Dim Res() As String
    Dim RN As Long
    Dim I As Long
    Dim Cur As String

    Src = Split(Replace(Replace(Code, vbCrLf, vbLf), vbCr, vbLf), vbLf)
    ReDim Res(0 To UBound(Src))

    I = 0
    Do While I <= UBound(Src)
        Cur = RTrimWs(Src(I))
        Do While EndsUnderscore(Cur) And I < UBound(Src)
            Cur = RTrimWs(RTrimWs(Left$(Cur, Len(Cur) - 1)) & " " & TrimWs(Src(I + 1)))
            I = I + 1
        Loop
        Res(RN) = Cur
        RN = RN + 1
        I = I + 1
    Loop

    ReDim Preserve Res(0 To RN - 1)
    JoinContinuations = Res

End Function

Private Function EndsUnderscore(ByVal S As String) As Boolean

    Dim N As Long

    N = Len(S)
    If N < 2 Then Exit Function
    If Mid$(S, N, 1) <> "_" Then Exit Function
    EndsUnderscore = (Mid$(S, N - 1, 1) = " " Or Mid$(S, N - 1, 1) = vbTab)

End Function

Private Function IsWordCh(ByVal C As String) As Boolean

    Dim A As Long

    If Len(C) = 0 Then Exit Function
    A = AscW(C)
    IsWordCh = (A >= 48 And A <= 57) Or (A >= 65 And A <= 90) Or (A >= 97 And A <= 122) Or A = 95 Or A > 127 Or A < 0

End Function

' ---------------------------------------------------------------------------
' 4 + 6. Limits, whitespace and procedure structure
' ---------------------------------------------------------------------------

Private Function StructuralProblems(ByVal Output As String, ByRef O As TFormatOptions) As String

    Dim Lines() As String
    Dim I As Long
    Dim Ln As String
    Dim Trimmed As String
    Dim Lead As String
    Dim Code As String
    Dim PrevCont As Boolean
    Dim Cont As Boolean
    Dim ContRun As Long
    Dim InProc As Boolean
    Dim ProcLine As Long
    Dim IsComment As Boolean
    Dim InlineEnd As Boolean
    Dim Probs As String
    Dim PCount As Long

    Lines = Split(Output, vbCrLf)

    For I = 0 To UBound(Lines)
        Ln = Lines(I)
        Trimmed = TrimWs(Ln)
        Lead = LeadingWs(Ln)

        If Len(Ln) > 0 Then
            If Right$(Ln, 1) = " " Or Right$(Ln, 1) = vbTab Then AddProb Probs, PCount, I + 1, "trailing whitespace"
        End If
        If Len(Ln) > 1023 Then AddProb Probs, PCount, I + 1, "longer than 1023 characters"

        If Not PrevCont And Len(Trimmed) > 0 Then
            If O.UseTabs Then
                If InStr(Lead, " ") > 0 Then AddProb Probs, PCount, I + 1, "spaces in indentation (UseTabs)"
            ElseIf InStr(Lead, vbTab) > 0 Then
                AddProb Probs, PCount, I + 1, "tab in indentation"
            ElseIf Len(Lead) Mod O.IndentSize <> 0 Then
                AddProb Probs, PCount, I + 1, "indent " & Len(Lead) & " is not a multiple of " & O.IndentSize
            End If
        End If

        Code = TrimWs(StripStringsAndComment(Ln))
        IsComment = (Left$(Trimmed, 1) = "'") Or Rx("^rem(\s|$)").Test(Trimmed)

        If Not PrevCont And Not IsComment And Len(Trimmed) > 0 Then
            InlineEnd = Rx("\bend\s+(sub|function|property)\b").Test(Code)
            If Rx("^(?:(?:public|private|friend|static|global)\s+)*(sub|function|property)\b").Test(Code) Or _
               Rx("^(?:(?:public|private|global)\s+)*(type|enum)\s+\w").Test(Code) Then
                If Len(Lead) > 0 Then AddProb Probs, PCount, I + 1, "procedure/type start is indented"
                If Not InlineEnd Then
                    If InProc Then AddProb Probs, PCount, I + 1, "starts while line " & ProcLine & " is still open"
                    InProc = True
                    ProcLine = I + 1
                End If
            ElseIf Rx("^end\s+(sub|function|property|type|enum)\b").Test(Code) Then
                If Len(Lead) > 0 Then AddProb Probs, PCount, I + 1, "'" & Code & "' is indented"
                If Not InProc Then AddProb Probs, PCount, I + 1, "'" & Code & "' without an open procedure"
                InProc = False
            ElseIf InProc And Len(Lead) = 0 Then
                If Left$(Code, 1) <> "#" And Not Rx("^[A-Za-z_]\w*:").Test(Code) And _
                   Not Rx("^\d+(\s|$)").Test(Code) And Not Rx("^attribute\s").Test(Code) Then
                    AddProb Probs, PCount, I + 1, "statement inside a procedure is at column 0: " & Left$(Trimmed, 40)
                End If
            End If
        End If

        Cont = EndsUnderscore(Ln)
        If Cont Then
            ContRun = ContRun + 1
            If ContRun > 24 Then AddProb Probs, PCount, I + 1, "more than 24 continuations in one statement"
        Else
            ContRun = 0
        End If
        PrevCont = Cont
    Next I

    If InProc Then AddProb Probs, PCount, UBound(Lines) + 1, "procedure from line " & ProcLine & " is never closed"

    StructuralProblems = Probs

End Function

Private Sub AddProb(ByRef Probs As String, ByRef PCount As Long, ByVal LineNo As Long, ByVal Text As String)

    PCount = PCount + 1
    If PCount <= 5 Then
        If Len(Probs) > 0 Then Probs = Probs & "; "
        Probs = Probs & "line " & LineNo & ": " & Text
    ElseIf PCount = 6 Then
        Probs = Probs & "; (more)"
    End If

End Sub

' Removes string contents and the trailing comment from one physical line.
Private Function StripStringsAndComment(ByVal Ln As String) As String

    Dim I As Long
    Dim Ch As String
    Dim InQ As Boolean
    Dim Res As String

    For I = 1 To Len(Ln)
        Ch = Mid$(Ln, I, 1)
        If InQ Then
            Res = Res & "~"
            If Ch = """" Then InQ = False
        ElseIf Ch = """" Then
            InQ = True
            Res = Res & "~"
        ElseIf Ch = "'" Then
            Exit For
        Else
            Res = Res & Ch
        End If
    Next I

    StripStringsAndComment = Res

End Function

' ---------------------------------------------------------------------------
' 5. Independent indentation oracle
' ---------------------------------------------------------------------------

Private Function OracleKind(ByVal Stmt As String) As String

    Dim T As String
    Dim Body As String
    Dim Parts() As String
    Dim Mt As Object
    Dim Lab As String

    T = TrimWs(Stmt)
    If Len(T) = 0 Then
        OracleKind = "blank"
        Exit Function
    End If
    OracleKind = "plain"

    If Left$(T, 1) = "#" Then OracleKind = "zero": Exit Function
    If Rx("^\d+(\s|:|$)").Test(T) Then OracleKind = "zero": Exit Function
    If Rx("^attribute\s").Test(T) Then OracleKind = "zero": Exit Function

    Set Mt = Rx("^([A-Za-z_]\w*):(?!=)").Execute(T)
    If Mt.Count > 0 Then
        Lab = LCase$(Mt(0).SubMatches(0))
        Select Case Lab
            Case "else", "case", "do", "loop", "next", "wend", "then"
            Case Else
                OracleKind = "zero"
                Exit Function
        End Select
    End If

    Parts = Split(T, ":")
    Body = Rx("^(?:(?:public|private|friend|static|global)\s+)*").Replace(T, "")

    If Rx("^(sub|function|property)\b").Test(Body) Then
        If Later(Parts, "^end\s+(sub|function|property)\b") Then OracleKind = "plain" Else OracleKind = "procOpen"
    ElseIf Rx("^(enum|type)\s+\w").Test(Body) Then
        OracleKind = "procOpen"
    ElseIf Rx("^end\s+(sub|function|property|enum|type)\b").Test(T) Then
        OracleKind = "procClose"
    ElseIf Rx("^if\b").Test(T) Then
        If Rx("\bthen\s*$").Test(T) Then OracleKind = "open"
    ElseIf Rx("^elseif\b").Test(T) Or Rx("^else\b").Test(T) Then
        OracleKind = "mid"
    ElseIf Rx("^end\s+if\b").Test(T) Or Rx("^end\s+with\b").Test(T) Then
        OracleKind = "close"
    ElseIf Rx("^select\s+case\b").Test(T) Then
        OracleKind = "selOpen"
    ElseIf Rx("^case\b").Test(T) Then
        OracleKind = "mid"
    ElseIf Rx("^end\s+select\b").Test(T) Then
        OracleKind = "selClose"
    ElseIf Rx("^for\b").Test(T) Then
        If Not Later(Parts, "^next\b") Then OracleKind = "open"
    ElseIf Rx("^next\b").Test(T) Then
        OracleKind = "close"
    ElseIf Rx("^do\b").Test(T) Then
        If Not Later(Parts, "^loop\b") Then OracleKind = "open"
    ElseIf Rx("^loop\b").Test(T) Then
        OracleKind = "close"
    ElseIf Rx("^while\b").Test(T) Then
        If Not Later(Parts, "^wend\b") Then OracleKind = "open"
    ElseIf Rx("^wend\b").Test(T) Then
        OracleKind = "close"
    ElseIf Rx("^with\b").Test(T) Then
        If Not Later(Parts, "^end\s+with\b") Then OracleKind = "open"
    End If

End Function

' True when any ":"-separated part after the first matches the pattern.
Private Function Later(ByRef Parts() As String, ByVal Pattern As String) As Boolean

    Dim I As Long

    For I = 1 To UBound(Parts)
        If Rx(Pattern).Test(TrimWs(Parts(I))) Then
            Later = True
            Exit Function
        End If
    Next I

End Function

Private Function IndentProblems(ByVal Output As String, ByRef O As TFormatOptions) As String

    Dim Lines() As String
    Dim I As Long
    Dim Last As Long
    Dim Ln As String
    Dim Trimmed As String
    Dim Lead As String
    Dim Actual As Long
    Dim Level As Long
    Dim Printed As Long
    Dim SelStep As Long
    Dim Kind As String
    Dim Code As String
    Dim IsComment As Boolean
    Dim CommentCont As Boolean
    Dim Probs As String
    Dim PCount As Long

    Lines = Split(Output, vbCrLf)
    If O.IndentCase Then SelStep = 2 Else SelStep = 1

    I = 0
    Do While I <= UBound(Lines)
        Ln = Lines(I)
        Trimmed = TrimWs(Ln)
        Lead = LeadingWs(Ln)
        If O.UseTabs Then Actual = Len(Lead) Else Actual = Len(Lead) \ O.IndentSize

        If CommentCont Then
            CommentCont = EndsUnderscore(Ln)
            I = I + 1
            GoTo NextLine
        End If
        If Len(Trimmed) = 0 Then
            I = I + 1
            GoTo NextLine
        End If

        IsComment = (Left$(Trimmed, 1) = "'") Or Rx("^rem(\s|$)").Test(Trimmed)
        If IsComment Then
            If Actual <> Level Then AddProb Probs, PCount, I + 1, "comment at level " & Actual & ", expected " & Level
            CommentCont = EndsUnderscore(Ln)
            I = I + 1
            GoTo NextLine
        End If

        ' the whole logical statement, continuation lines joined
        Code = TrimWs(StripStringsAndComment(Ln))
        Last = I
        Do While EndsUnderscore(Lines(Last)) And Last < UBound(Lines)
            If Len(TrimWs(Lines(Last + 1))) = 0 Then Exit Do
            If Right$(Code, 1) = "_" Then Code = RTrimWs(Left$(Code, Len(Code) - 1))
            Code = Code & " " & TrimWs(StripStringsAndComment(Lines(Last + 1)))
            Last = Last + 1
        Loop

        Kind = OracleKind(Code)
        Printed = Level
        Select Case Kind
            Case "close"
                Level = Level - 1
                If Level < 0 Then Level = 0
                Printed = Level
            Case "selClose"
                Level = Level - SelStep
                If Level < 0 Then Level = 0
                Printed = Level
            Case "procOpen", "procClose"
                Level = 0
                Printed = 0
            Case "mid"
                Printed = Level - 1
                If Printed < 0 Then Printed = 0
            Case "zero"
                Printed = 0
        End Select

        If Actual <> Printed Then
            AddProb Probs, PCount, I + 1, Kind & " at level " & Actual & ", expected " & Printed & ": " & Left$(Trimmed, 40)
        End If

        If Kind = "open" Or Kind = "procOpen" Then
            Level = Level + 1
        ElseIf Kind = "selOpen" Then
            Level = Level + SelStep
        End If

        I = Last + 1
NextLine:
    Loop

    IndentProblems = Probs

End Function

' ---------------------------------------------------------------------------
' 3. Input variants that must not change the output
' ---------------------------------------------------------------------------

Private Function MutationName(ByVal M As Long) As String

    Select Case M
        Case 1: MutationName = "no indentation"
        Case 2: MutationName = "tab indentation"
        Case 3: MutationName = "random deep indentation"
        Case 4: MutationName = "trailing spaces"
        Case 5: MutationName = "CR line endings"
        Case 6: MutationName = "CRLF line endings"
    End Select

End Function

Private Function Mutate(ByVal Code As String, ByVal M As Long) As String

    Dim Lines() As String
    Dim I As Long
    Dim T As String
    Dim N As Long
    Dim Norm As String

    Norm = Replace(Replace(Code, vbCrLf, vbLf), vbCr, vbLf)
    Lines = Split(Norm, vbLf)

    For I = 0 To UBound(Lines)
        T = Lines(I)
        Select Case M
            Case 1
                Lines(I) = LTrimWs(T)
            Case 2
                N = 0
                Do While Mid$(T, N * 4 + 1, 4) = "    "
                    N = N + 1
                Loop
                Lines(I) = String$(N, vbTab) & Mid$(T, N * 4 + 1)
            Case 3
                If Len(TrimWs(T)) > 0 Then Lines(I) = Space$((I * 7) Mod 23) & TrimWs(T)
            Case 4
                If I Mod 2 = 1 Then Lines(I) = T & " " & vbTab & " " Else Lines(I) = T & "  "
        End Select
    Next I

    Mutate = Join(Lines, vbLf)
    If M = 5 Then Mutate = Replace(Norm, vbLf, vbCr)
    If M = 6 Then Mutate = Replace(Norm, vbLf, vbCrLf)

End Function

' ===========================================================================
' Reporting and helpers
' ===========================================================================

Private Sub AddFail(ByRef Report As String, ByRef Fails As Long, ByVal Text As String)

    Fails = Fails + 1
    Report = Report & "FAIL  " & Text & vbCrLf

End Sub

Private Sub Finish(ByVal Title As String, ByVal Report As String, ByVal Checks As Long, ByVal Fails As Long)

    Dim Path As String
    Dim F As Integer
    Dim Summary As String
    Dim Parts() As String
    Dim I As Long

    Summary = Title & ": " & Checks & " checks, " & Fails & " failures."

    Debug.Print String$(60, "-")
    If Len(Report) > 0 Then
        Parts = Split(Report, vbCrLf)
        For I = 0 To UBound(Parts)
            If I >= 60 Then
                Debug.Print "... (see the report file for the rest)"
                Exit For
            End If
            Debug.Print Parts(I)
        Next I
    End If
    Debug.Print Summary

    If Fails > 0 Then
        Path = Environ$("TEMP") & "\VBAFormatterTestReport.txt"
        On Error Resume Next
        F = FreeFile
        Open Path For Output As #F
        Print #F, Summary
        Print #F, ""
        Print #F, Report
        Close #F
        On Error GoTo 0
        MsgBox Summary & vbCrLf & vbCrLf & "Details: Immediate window (Ctrl+G) and" & vbCrLf & Path, vbExclamation, "Formatter tests"
    Else
        MsgBox Summary, vbInformation, "Formatter tests"
    End If

End Sub

Private Function FirstDiff(ByVal A As String, ByVal B As String) As String

    Dim X() As String
    Dim Y() As String
    Dim I As Long
    Dim N As Long

    X = Split(A, vbCrLf)
    Y = Split(B, vbCrLf)
    N = UBound(X)
    If UBound(Y) < N Then N = UBound(Y)

    For I = 0 To N
        If X(I) <> Y(I) Then
            FirstDiff = "line " & (I + 1) & ": [" & X(I) & "] vs [" & Y(I) & "]"
            Exit Function
        End If
    Next I

    FirstDiff = "different number of lines (" & (UBound(X) + 1) & " vs " & (UBound(Y) + 1) & ")"

End Function

Private Function FirstTokenDiff(ByVal KeyA As String, ByVal KeyB As String) As String

    Dim X() As String
    Dim Y() As String
    Dim I As Long
    Dim N As Long

    X = Split(KeyA, Chr$(1))
    Y = Split(KeyB, Chr$(1))
    N = UBound(X)
    If UBound(Y) < N Then N = UBound(Y)

    For I = 0 To N
        If X(I) <> Y(I) Then
            FirstTokenDiff = "token " & (I + 1) & ": [" & Replace(X(I), vbLf, "<EOL>") & "] became [" & Replace(Y(I), vbLf, "<EOL>") & "]"
            Exit Function
        End If
    Next I

    FirstTokenDiff = "token count " & (UBound(X) + 1) & " became " & (UBound(Y) + 1)

End Function

' Strips the VERSION ... BEGIN ... END header of exported .cls / .frm files.
Private Function ExtractCode(ByVal Text As String) As String

    Dim Lines() As String
    Dim I As Long
    Dim Depth As Long
    Dim T As String
    Dim Res() As String
    Dim RN As Long

    Lines = Split(Replace(Replace(Text, vbCrLf, vbLf), vbCr, vbLf), vbLf)
    I = 0

    If UBound(Lines) >= 0 Then
        If UCase$(Left$(Lines(0), 8)) = "VERSION " Then
            I = 1
            Do While I <= UBound(Lines)
                T = UCase$(TrimWs(Lines(I)))
                If T = "BEGIN" Or Left$(T, 6) = "BEGIN " Then
                    Depth = Depth + 1
                ElseIf T = "END" Then
                    Depth = Depth - 1
                    If Depth <= 0 Then
                        I = I + 1
                        Exit Do
                    End If
                End If
                I = I + 1
            Loop
        End If
    End If

    If I > UBound(Lines) Then Exit Function
    ReDim Res(0 To UBound(Lines) - I)
    Do While I <= UBound(Lines)
        Res(RN) = Lines(I)
        RN = RN + 1
        I = I + 1
    Loop

    ExtractCode = Join(Res, vbLf)

End Function

Private Function ReadFile(ByVal Path As String) As String

    Dim F As Integer
    Dim Ln As String
    Dim Buf() As String
    Dim N As Long

    ReDim Buf(1 To 256)
    F = FreeFile
    Open Path For Input As #F
    Do While Not EOF(F)
        Line Input #F, Ln
        N = N + 1
        If N > UBound(Buf) Then ReDim Preserve Buf(1 To UBound(Buf) * 2)
        Buf(N) = Ln
    Loop
    Close #F

    If N = 0 Then Exit Function
    ReDim Preserve Buf(1 To N)
    ReadFile = Join(Buf, vbLf)

End Function

Private Function PickFolder(ByVal Title As String, ByVal StartFolder As String) As String

    Dim Dlg As Object

    On Error GoTo NoDialog
    Set Dlg = Application.FileDialog(4)
    Dlg.Title = Title
    If Len(StartFolder) > 0 Then
        If Len(Dir$(StartFolder, vbDirectory)) > 0 Then Dlg.InitialFileName = StartFolder & "\"
    End If
    If Dlg.Show <> -1 Then Exit Function
    PickFolder = Dlg.SelectedItems(1)
    Exit Function

NoDialog:
    PickFolder = StartFolder

End Function

' Compiled, cached regular expression (case-insensitive).
Private Function Rx(ByVal Pattern As String) As Object

    Dim R As Object

    If mCache Is Nothing Then Set mCache = CreateObject("Scripting.Dictionary")

    If mCache.Exists(Pattern) Then
        Set Rx = mCache(Pattern)
    Else
        Set R = CreateObject("VBScript.RegExp")
        R.Pattern = Pattern
        R.IgnoreCase = True
        R.Global = False
        Set mCache(Pattern) = R
        Set Rx = R
    End If

End Function

Private Function LeadingWs(ByVal S As String) As String

    Dim I As Long

    I = 1
    Do While I <= Len(S)
        If Mid$(S, I, 1) = " " Or Mid$(S, I, 1) = vbTab Then I = I + 1 Else Exit Do
    Loop

    LeadingWs = Left$(S, I - 1)

End Function

Private Function LTrimWs(ByVal S As String) As String

    LTrimWs = Mid$(S, Len(LeadingWs(S)) + 1)

End Function

Private Function RTrimWs(ByVal S As String) As String

    Dim N As Long

    N = Len(S)
    Do While N > 0
        If Mid$(S, N, 1) = " " Or Mid$(S, N, 1) = vbTab Or Mid$(S, N, 1) = vbCr Or Mid$(S, N, 1) = vbLf Then
            N = N - 1
        Else
            Exit Do
        End If
    Loop

    RTrimWs = Left$(S, N)

End Function

Private Function TrimWs(ByVal S As String) As String

    TrimWs = LTrimWs(RTrimWs(S))

End Function
