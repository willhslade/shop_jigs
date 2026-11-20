Attribute VB_Name = "modFormatterCore"
Option Explicit

' ---------------------------------------------------------------------------
' VBA formatter engine. Pure string in / string out - no VBIDE dependency.
'
' Pipeline
'   1. ParseStatements : split into logical statements (joins " _" continuations),
'                        mask strings, strip comments, classify, compute indent.
'   2. LayoutPass      : blank-line rules (collapse, between procedures, after declarations).
'   3. Align passes    : optional "As" and trailing-comment alignment.
'   4. Render          : produce text.
'
' Known limits: statements chained with ":" are classified by their first
' statement only (inline Do..Loop / For..Next / With..End With are detected);
' "Else If" on one line is treated as Else; #If blocks are left-aligned and
' do not affect indentation.
' ---------------------------------------------------------------------------

Private Const K_BLANK As Long = 0
Private Const K_COMMENT As Long = 1
Private Const K_CODE As Long = 2
Private Const K_DECL As Long = 3
Private Const K_OPEN As Long = 4
Private Const K_CLOSE As Long = 5
Private Const K_MID As Long = 6
Private Const K_PROC_OPEN As Long = 7
Private Const K_PROC_CLOSE As Long = 8
Private Const K_SELECT_OPEN As Long = 9
Private Const K_SELECT_CLOSE As Long = 10
Private Const K_DIRECTIVE As Long = 11
Private Const K_LABEL As Long = 12
Private Const K_RAW As Long = 13

Private Const MASK_CHAR As String = "~"

' VBA allows at most 24 line continuations (25 physical lines) per statement
' and 1023 characters per physical line.
Private Const MAX_PHYS_LINES As Long = 25
Private Const MAX_PHYS_CHARS As Long = 1000

' In VBA a comment that ends in " _" continues onto the next physical line, so that line is
' comment text, not code. Set to False if your VBE version does not behave this way.
Private Const COMMENTS_CONTINUE As Boolean = True

Private Type TOutLine
    Kind As Long
    Level As Long
    Body As String          ' code without indent/comment; multi-line statements hold full indented text
    Gap As String           ' whitespace between code and trailing comment
    Comment As String       ' trailing comment, or the whole text of a comment-only line
    Multi As Boolean        ' statement spans several physical lines (excluded from alignment)
    VarDecl As Boolean      ' eligible for "As" alignment
End Type

' ===========================================================================
' Public entry points
' ===========================================================================

Public Function FormatVbaCode(ByVal CodeText As String) As String

    Dim Opt As TFormatOptions

    Opt = LoadOptions()
    FormatVbaCode = FormatWithOptions(CodeText, Opt)

End Function

' InitialLevel : -1 = start at level 0; >= 0 = start at this indent level (fragments)
' IsFragment   : True for a procedure/selection (never inserts Option Explicit)
Public Function FormatWithOptions(ByVal CodeText As String, ByRef Opt As TFormatOptions, _
                                  Optional ByVal InitialLevel As Long = -1, _
                                  Optional ByVal IsFragment As Boolean = False) As String

    Dim Src() As String
    Dim P() As TOutLine
    Dim R() As TOutLine
    Dim Rendered() As String
    Dim PN As Long
    Dim RN As Long
    Dim I As Long

    CodeText = Replace(Replace(CodeText, vbCrLf, vbLf), vbCr, vbLf)
    If Len(TrimAll(CodeText)) = 0 Then Exit Function

    Src = Split(CodeText, vbLf)

    ReDim P(1 To 64)
    ReDim R(1 To 64)

    ParseStatements Src, Opt, InitialLevel, P, PN

    If Opt.AddOptionExplicit And Not IsFragment Then InsertOptionExplicit P, PN

    LayoutPass P, PN, Opt, R, RN

    If RN = 0 Then Exit Function

    If Opt.AlignAs Then AlignDeclarations R, RN
    If Opt.AlignComments Then AlignTrailingComments R, RN, Opt

    ReDim Rendered(1 To RN)
    For I = 1 To RN
        Rendered(I) = RenderLine(R(I), Opt)
    Next I

    FormatWithOptions = Join(Rendered, vbCrLf)

End Function

' ===========================================================================
' Pass 1: parse, classify, indent
' ===========================================================================

Private Sub ParseStatements(ByRef Src() As String, ByRef Opt As TFormatOptions, _
                            ByVal InitialLevel As Long, _
                            ByRef P() As TOutLine, ByRef PN As Long)

    Dim I As Long
    Dim K As Long
    Dim NPhys As Long
    Dim Level As Long
    Dim PrintLevel As Long
    Dim SelStep As Long
    Dim Kind As Long
    Dim FirstStmt As Boolean
    Dim Multi As Boolean
    Dim VarDecl As Boolean
    Dim IsSig As Boolean
    Dim T As String
    Dim Cmt As String
    Dim Gap As String
    Dim JoinM As String
    Dim Piece As String
    Dim Body As String
    Dim C2 As String
    Dim M2 As String
    Dim IndentTxt As String
    Dim Code() As String
    Dim Mask() As String
    Dim Ln As TOutLine
    Dim Blank0 As TOutLine

    If Opt.IndentCase Then SelStep = 2 Else SelStep = 1
    If InitialLevel > 0 Then Level = InitialLevel
    FirstStmt = True

    I = LBound(Src)
    Do While I <= UBound(Src)

        T = TrimAll(Src(I))
        Ln = Blank0

        If Len(T) = 0 Then

            Ln.Kind = K_BLANK
            PushLine P, PN, Ln
            I = I + 1

        ElseIf IsWholeLineComment(T) Then

            Ln.Kind = K_COMMENT
            Ln.Level = Level
            Ln.Comment = T
            PushLine P, PN, Ln
            I = I + 1
            I = ConsumeCommentContinuation(Src, I, T, Level, P, PN)

        ElseIf IsAttributeLine(T) Then

            Ln.Kind = K_RAW
            Ln.Body = T
            PushLine P, PN, Ln
            I = I + 1

        Else

            NPhys = ReadStatement(Src, I, Code, Mask, Cmt, Gap)

            ' Logical statement, masked, for classification
            JoinM = ""
            For K = 1 To NPhys
                If K < NPhys Then
                    ContBody Code(K), Mask(K), C2, M2
                Else
                    C2 = Code(K)
                    M2 = Mask(K)
                End If
                If Len(M2) > 0 Then
                    If Len(JoinM) > 0 Then JoinM = JoinM & " "
                    JoinM = JoinM & M2
                End If
            Next K

            Kind = ClassifyStatement(JoinM, VarDecl, IsSig)

            If FirstStmt Then
                FirstStmt = False
                If InitialLevel >= 0 Then
                    Select Case Kind
                        Case K_CLOSE, K_MID
                            Level = InitialLevel + 1
                        Case K_SELECT_CLOSE
                            Level = InitialLevel + SelStep
                        Case Else
                            Level = InitialLevel
                    End Select
                End If
            End If

            ' Indent before the line
            PrintLevel = Level
            Select Case Kind
                Case K_CLOSE
                    Level = Level - 1
                    If Level < 0 Then Level = 0
                    PrintLevel = Level
                Case K_SELECT_CLOSE
                    Level = Level - SelStep
                    If Level < 0 Then Level = 0
                    PrintLevel = Level
                Case K_PROC_CLOSE, K_PROC_OPEN
                    Level = 0
                    PrintLevel = 0
                Case K_MID
                    PrintLevel = Level - 1
                    If PrintLevel < 0 Then PrintLevel = 0
                Case K_DIRECTIVE, K_LABEL
                    PrintLevel = 0
            End Select

            ' Build the text
            IndentTxt = IndentString(PrintLevel, Opt)
            Multi = False

            If IsSig And Opt.SignatureWrap <> 0 Then
                Body = BuildSignature(Code, Mask, NPhys, PrintLevel, Len(Cmt), Opt, Multi)
            ElseIf NPhys = 1 Then
                Body = SpaceCode(Code(1), Mask(1), Opt)
            Else
                Body = ""
                For K = 1 To NPhys
                    If K < NPhys Then
                        ContBody Code(K), Mask(K), C2, M2
                        Piece = SpaceCode(C2, M2, Opt) & " _"
                    Else
                        Piece = SpaceCode(Code(K), Mask(K), Opt)
                    End If
                    If K = 1 Then
                        Piece = IndentTxt & Piece
                    Else
                        Piece = IndentString(PrintLevel + Opt.ContinuationLevels, Opt) & Piece
                    End If
                    If K > 1 Then Body = Body & vbCrLf
                    Body = Body & Piece
                Next K
                Multi = True
            End If

            Ln.Kind = Kind
            Ln.Level = PrintLevel
            Ln.Body = Body
            Ln.Gap = Gap
            Ln.Comment = Cmt
            Ln.Multi = Multi
            Ln.VarDecl = (VarDecl And Not Multi)
            PushLine P, PN, Ln

            ' Indent after the line
            Select Case Kind
                Case K_OPEN, K_PROC_OPEN
                    Level = Level + 1
                Case K_SELECT_OPEN
                    Level = Level + SelStep
            End Select

            I = I + NPhys

            If Len(Cmt) > 0 Then
                I = ConsumeCommentContinuation(Src, I, Cmt, PrintLevel + Opt.ContinuationLevels, P, PN)
            End If

        End If

    Loop

End Sub

Private Function EndsCommentContinuation(ByVal CommentText As String) As Boolean

    Dim N As Long
    Dim B As String

    If Not COMMENTS_CONTINUE Then Exit Function
    N = Len(CommentText)
    If N < 2 Then Exit Function
    If Mid$(CommentText, N, 1) <> "_" Then Exit Function
    B = Mid$(CommentText, N - 1, 1)
    EndsCommentContinuation = (B = " " Or B = vbTab)

End Function

' Lines that follow a comment ending in " _" are part of that comment: keep them as comment
' lines (re-indented) instead of classifying them as code. Returns the next unread index.
Private Function ConsumeCommentContinuation(ByRef Src() As String, ByVal I As Long, _
                                            ByVal CommentText As String, ByVal Level As Long, _
                                            ByRef P() As TOutLine, ByRef PN As Long) As Long

    Dim Cur As String
    Dim T As String
    Dim Ln As TOutLine

    Cur = CommentText
    Do While EndsCommentContinuation(Cur) And I <= UBound(Src)
        T = TrimAll(Src(I))
        If Len(T) = 0 Then Exit Do
        Ln.Kind = K_COMMENT
        Ln.Level = Level
        Ln.Comment = T
        PushLine P, PN, Ln
        Cur = T
        I = I + 1
    Loop

    ConsumeCommentContinuation = I

End Function

' Reads one logical statement starting at StartIdx. Code()/Mask() hold the trimmed
' code of each physical line (comment removed); Mask has string contents replaced.
Private Function ReadStatement(ByRef Src() As String, ByVal StartIdx As Long, _
                               ByRef Code() As String, ByRef Mask() As String, _
                               ByRef Cmt As String, ByRef Gap As String) As Long

    Dim N As Long
    Dim Idx As Long
    Dim C As String
    Dim M As String
    Dim K As String
    Dim G As String

    Idx = StartIdx
    Do
        SplitLine TrimAll(Src(Idx)), C, M, K, G
        N = N + 1
        If N = 1 Then
            ReDim Code(1 To 1)
            ReDim Mask(1 To 1)
        Else
            ReDim Preserve Code(1 To N)
            ReDim Preserve Mask(1 To N)
        End If
        Code(N) = C
        Mask(N) = M
        Cmt = K
        Gap = G

        If Not EndsWithContinuation(M) Then Exit Do
        If Idx >= UBound(Src) Then Exit Do
        If Len(TrimAll(Src(Idx + 1))) = 0 Then Exit Do
        Idx = Idx + 1
    Loop

    ReadStatement = N

End Function

' Splits a trimmed physical line into code, masked code, trailing comment and the
' whitespace between code and comment. Quotes are honoured ("" inside a string).
Private Sub SplitLine(ByVal T As String, ByRef Code As String, ByRef Masked As String, _
                      ByRef Cmt As String, ByRef Gap As String)

    Dim I As Long
    Dim N As Long
    Dim Keep As Long
    Dim Ch As String
    Dim M As String
    Dim InQuote As Boolean

    N = Len(T)
    M = T
    Cmt = ""
    Gap = ""
    Code = T

    For I = 1 To N
        Ch = Mid$(T, I, 1)
        If InQuote Then
            Mid$(M, I, 1) = MASK_CHAR
            If Ch = """" Then InQuote = False
        ElseIf Ch = """" Then
            InQuote = True
            Mid$(M, I, 1) = MASK_CHAR
        ElseIf Ch = "'" Then
            Cmt = RTrimWs(Mid$(T, I))
            Code = Left$(T, I - 1)
            M = Left$(M, I - 1)
            Keep = Len(RTrimWs(Code))
            Gap = Mid$(Code, Keep + 1)
            Code = Left$(Code, Keep)
            M = Left$(M, Keep)
            Exit For
        End If
    Next I

    Masked = M

End Sub

Private Function EndsWithContinuation(ByVal M As String) As Boolean

    Dim N As Long
    Dim B As String

    N = Len(M)
    If N < 2 Then Exit Function
    If Mid$(M, N, 1) <> "_" Then Exit Function
    B = Mid$(M, N - 1, 1)
    EndsWithContinuation = (B = " " Or B = vbTab)

End Function

' Removes the trailing " _" (and the whitespace before it) from a continued line.
Private Sub ContBody(ByVal Code As String, ByVal Masked As String, _
                     ByRef OutCode As String, ByRef OutMask As String)

    Dim L As Long

    L = Len(Code) - 1
    If L < 0 Then L = 0
    Do While L > 0
        If Mid$(Code, L, 1) = " " Or Mid$(Code, L, 1) = vbTab Then
            L = L - 1
        Else
            Exit Do
        End If
    Loop

    OutCode = Left$(Code, L)
    OutMask = Left$(Masked, L)

End Sub

' ===========================================================================
' Classification
' ===========================================================================

Private Function ClassifyStatement(ByVal M As String, ByRef VarDecl As Boolean, _
                                   ByRef IsSig As Boolean) As Long

    Dim W() As String
    Dim N As Long
    Dim I As Long
    Dim Kw As String

    VarDecl = False
    IsSig = False
    ClassifyStatement = K_CODE

    N = GetWords(M, W)
    If N = 0 Then Exit Function

    If Left$(W(1), 1) = "#" Then
        ClassifyStatement = K_DIRECTIVE
        Exit Function
    End If

    If IsLabelStatement(M, W(1)) Then
        ClassifyStatement = K_LABEL
        Exit Function
    End If

    ' Skip Public/Private/Friend/Global/Static in front of Sub/Function/...
    I = 1
    Do While I < N
        If IsModifierWord(W(I)) Then
            I = I + 1
        Else
            Exit Do
        End If
    Loop
    Kw = W(I)

    Select Case Kw

        Case "SUB", "FUNCTION", "PROPERTY"
            If InlineCloser(M, "END", Kw) Then
                ClassifyStatement = K_CODE
            Else
                ClassifyStatement = K_PROC_OPEN
                IsSig = True
            End If

        Case "ENUM", "TYPE"
            ClassifyStatement = K_PROC_OPEN

        Case "DECLARE"
            ClassifyStatement = K_DECL
            IsSig = True

        Case "IF"
            If IsBlockIf(M) Then ClassifyStatement = K_OPEN

        Case "ELSEIF", "ELSE", "CASE"
            ClassifyStatement = K_MID

        Case "SELECT"
            If W(2) = "CASE" Then ClassifyStatement = K_SELECT_OPEN

        Case "FOR"
            If Not InlineCloser(M, "NEXT", "") Then ClassifyStatement = K_OPEN

        Case "DO"
            If Not InlineCloser(M, "LOOP", "") Then ClassifyStatement = K_OPEN

        Case "WHILE"
            If Not InlineCloser(M, "WEND", "") Then ClassifyStatement = K_OPEN

        Case "WITH"
            If Not InlineCloser(M, "END", "WITH") Then ClassifyStatement = K_OPEN

        Case "NEXT", "LOOP", "WEND"
            ClassifyStatement = K_CLOSE

        Case "END"
            Select Case W(2)
                Case "IF", "WITH"
                    ClassifyStatement = K_CLOSE
                Case "SELECT"
                    ClassifyStatement = K_SELECT_CLOSE
                Case "SUB", "FUNCTION", "PROPERTY", "TYPE", "ENUM"
                    ClassifyStatement = K_PROC_CLOSE
            End Select

        Case Else
            Select Case W(1)
                Case "DIM", "CONST", "STATIC", "PUBLIC", "PRIVATE", "GLOBAL", "FRIEND", _
                     "OPTION", "IMPLEMENTS", "EVENT"
                    ClassifyStatement = K_DECL
                    Select Case W(1)
                        Case "DIM", "CONST", "STATIC", "PUBLIC", "PRIVATE", "GLOBAL"
                            VarDecl = (W(2) <> "EVENT")
                    End Select
            End Select

    End Select

End Function

Private Function IsModifierWord(ByVal W As String) As Boolean

    Select Case W
        Case "PUBLIC", "PRIVATE", "FRIEND", "GLOBAL", "STATIC"
            IsModifierWord = True
    End Select

End Function

' First whole word (up to four) of a masked statement, upper-cased. A "#" prefix stays
' attached (#If). Any other non-word character becomes a one-character word.
Private Function GetWords(ByVal S As String, ByRef W() As String) As Long

    Dim I As Long
    Dim L As Long
    Dim St As Long
    Dim Cnt As Long
    Dim Ch As String

    ReDim W(1 To 4)
    L = Len(S)
    I = 1

    Do While I <= L And Cnt < 4
        Ch = Mid$(S, I, 1)
        If Ch = " " Or Ch = vbTab Then
            I = I + 1
        Else
            St = I
            If Ch = "#" Then
                I = I + 1
                Do While I <= L
                    If IsWordChar(Mid$(S, I, 1)) Then I = I + 1 Else Exit Do
                Loop
            ElseIf IsWordChar(Ch) Then
                Do While I <= L
                    If IsWordChar(Mid$(S, I, 1)) Then I = I + 1 Else Exit Do
                Loop
            Else
                I = I + 1
            End If
            Cnt = Cnt + 1
            W(Cnt) = UCase$(Mid$(S, St, I - St))
        End If
    Loop

    GetWords = Cnt

End Function

Private Function IsWordChar(ByVal C As String) As Boolean

    If Len(C) = 0 Then Exit Function
    Select Case AscW(C)
        Case 48 To 57, 65 To 90, 97 To 122, 95, Is > 127, Is < 0
            IsWordChar = True
    End Select

End Function

Private Function IsLabelStatement(ByVal M As String, ByVal W1 As String) As Boolean

    Dim L As Long

    If Len(W1) = 0 Then Exit Function

    ' a statement cannot start with a digit: it is a numeric line label
    If Not (W1 Like "*[!0-9]*") Then
        IsLabelStatement = True
        Exit Function
    End If

    If Not (W1 Like "[A-Z]*") Then Exit Function

    Select Case W1
        Case "ELSE", "DO", "LOOP", "NEXT", "WEND", "THEN"
            Exit Function
    End Select

    L = Len(W1)
    If Mid$(M, L + 1, 1) <> ":" Then Exit Function
    If Mid$(M, L + 2, 1) = "=" Then Exit Function

    IsLabelStatement = True

End Function

' True for "If ... Then" with nothing after Then (a block If). Input is masked.
Private Function IsBlockIf(ByVal M As String) As Boolean

    Dim U As String
    Dim P As Long

    U = UCase$(M)
    P = FindWord(U, "THEN")
    If P = 0 Then Exit Function

    IsBlockIf = (Len(TrimAll(Mid$(U, P + 4))) = 0)

End Function

' True when a ":"-separated statement after the first one starts with W1 [W2],
' e.g. "For i = 1 To 3: x = x + i: Next".
Private Function InlineCloser(ByVal M As String, ByVal W1 As String, ByVal W2 As String) As Boolean

    Dim P As Long
    Dim W() As String
    Dim Dummy As Long

    P = InStr(1, M, ":")
    Do While P > 0
        If Mid$(M, P + 1, 1) = "=" Then
            P = P + 1
        Else
            Dummy = GetWords(Mid$(M, P + 1), W)
            If W(1) = W1 Then
                If Len(W2) = 0 Then
                    InlineCloser = True
                    Exit Function
                ElseIf W(2) = W2 Then
                    InlineCloser = True
                    Exit Function
                End If
            End If
        End If
        P = InStr(P + 1, M, ":")
    Loop

End Function

Private Function FindWord(ByVal U As String, ByVal Word As String, Optional ByVal Start As Long = 1) As Long

    Dim P As Long
    Dim L As Long
    Dim OkBefore As Boolean

    L = Len(Word)
    P = InStr(Start, U, Word, vbBinaryCompare)

    Do While P > 0
        OkBefore = True
        If P > 1 Then OkBefore = Not IsWordChar(Mid$(U, P - 1, 1))
        If OkBefore Then
            If Not IsWordChar(Mid$(U, P + L, 1)) Then
                FindWord = P
                Exit Function
            End If
        End If
        P = InStr(P + 1, U, Word, vbBinaryCompare)
    Loop

End Function

' ===========================================================================
' Text builders
' ===========================================================================

' One space around = <> <= >= < > ; ", " after commas; never inside strings.
' Code and Masked have identical length (Masked has string contents replaced).
Private Function SpaceCode(ByVal Code As String, ByVal Masked As String, _
                           ByRef Opt As TFormatOptions) As String

    Dim I As Long
    Dim N As Long
    Dim C As String
    Dim Nx As String
    Dim Op As String
    Dim Res As String

    If Not (Opt.SpaceOperators Or Opt.SpaceCommas) Then
        SpaceCode = Code
        Exit Function
    End If

    N = Len(Code)
    I = 1
    Do While I <= N
        C = Mid$(Masked, I, 1)

        If C = MASK_CHAR Then

            Res = Res & Mid$(Code, I, 1)
            I = I + 1

        ElseIf C = "," And Opt.SpaceCommas Then

            Res = RTrimWs(Res) & ","
            I = I + 1
            Do While I <= N
                Nx = Mid$(Masked, I, 1)
                If Nx = " " Or Nx = vbTab Then I = I + 1 Else Exit Do
            Loop
            If I <= N Then
                If Mid$(Masked, I, 1) <> ")" Then Res = Res & " "
            End If

        ElseIf C = ":" And Mid$(Masked, I + 1, 1) = "=" Then

            Res = Res & ":="
            I = I + 2

        ElseIf Opt.SpaceOperators And (C = "<" Or C = ">" Or C = "=") Then

            Op = C
            Nx = Mid$(Masked, I + 1, 1)
            If (Nx = "<" Or Nx = ">" Or Nx = "=") And Nx <> C Then Op = C & Nx
            I = I + Len(Op)
            Res = RTrimWs(Res) & " " & Op & " "
            Do While I <= N
                Nx = Mid$(Masked, I, 1)
                If Nx = " " Or Nx = vbTab Then I = I + 1 Else Exit Do
            Loop

        Else

            Res = Res & Mid$(Code, I, 1)
            I = I + 1

        End If
    Loop

    SpaceCode = RTrimWs(Res)

End Function

' Sub/Function/Property/Declare line. Joins continuation lines, then either
' returns one line or a wrapped version (Multi = True, indentation embedded).
Private Function BuildSignature(ByRef Code() As String, ByRef Mask() As String, ByVal NPhys As Long, _
                                ByVal PrintLevel As Long, ByVal CmtLen As Long, _
                                ByRef Opt As TFormatOptions, ByRef Multi As Boolean) As String

    Dim J As String
    Dim JM As String
    Dim C2 As String
    Dim M2 As String
    Dim K As Long
    Dim I As Long
    Dim P1 As Long
    Dim P2 As Long
    Dim Cap As Long
    Dim PC As Long
    Dim Depth As Long
    Dim St As Long
    Dim Ch As String
    Dim Inner As String
    Dim Params() As String
    Dim Head As String
    Dim Tail As String
    Dim OneLine As String
    Dim LineWidth As Long
    Dim DoWrap As Boolean
    Dim OnePer As Boolean
    Dim Limit As Long
    Dim LineCount As Long
    Dim Res As String

    Multi = False

    For K = 1 To NPhys
        If K < NPhys Then
            ContBody Code(K), Mask(K), C2, M2
        Else
            C2 = Code(K)
            M2 = Mask(K)
        End If
        If Len(C2) > 0 Then
            If Len(J) > 0 Then
                J = J & " "
                JM = JM & " "
            End If
            J = J & C2
            JM = JM & M2
        End If
    Next K

    P1 = InStr(JM, "(")
    If P1 > 0 Then P2 = MatchParen(JM, P1)
    If P1 = 0 Or P2 = 0 Then
        BuildSignature = SpaceCode(J, JM, Opt)
        Exit Function
    End If

    Inner = Mid$(JM, P1 + 1, P2 - P1 - 1)
    If Len(TrimAll(Inner)) = 0 Then
        BuildSignature = SpaceCode(J, JM, Opt)
        Exit Function
    End If

    ' Split parameters on top-level commas
    Cap = (Len(Inner) - Len(Replace(Inner, ",", ""))) + 1
    ReDim Params(1 To Cap)
    St = P1 + 1
    For I = P1 + 1 To P2 - 1
        Ch = Mid$(JM, I, 1)
        If Ch = "(" Then
            Depth = Depth + 1
        ElseIf Ch = ")" Then
            Depth = Depth - 1
        ElseIf Ch = "," And Depth = 0 Then
            PC = PC + 1
            Params(PC) = TrimAll(SpaceCode(Mid$(J, St, I - St), Mid$(JM, St, I - St), Opt))
            St = I + 1
        End If
    Next I
    PC = PC + 1
    Params(PC) = TrimAll(SpaceCode(Mid$(J, St, P2 - St), Mid$(JM, St, P2 - St), Opt))

    Head = SpaceCode(Left$(J, P1), Left$(JM, P1), Opt)
    Tail = TrimAll(SpaceCode(Mid$(J, P2 + 1), Mid$(JM, P2 + 1), Opt))

    OneLine = Head
    For I = 1 To PC
        If I > 1 Then OneLine = OneLine & ", "
        OneLine = OneLine & Params(I)
    Next I
    OneLine = OneLine & ")"
    If Len(Tail) > 0 Then OneLine = OneLine & " " & Tail

    LineWidth = PrintLevel * Opt.IndentSize + Len(OneLine)
    If CmtLen > 0 Then LineWidth = LineWidth + 1 + CmtLen

    Select Case Opt.SignatureWrap
        Case 1
            DoWrap = (Opt.MaxLineLength > 0 And LineWidth > Opt.MaxLineLength)
        Case 2
            DoWrap = (PC >= 2)
    End Select

    ' A single line that would break VBA's 1023 character limit must wrap
    If PrintLevel * Opt.IndentSize + Len(OneLine) > MAX_PHYS_CHARS Then DoWrap = True

    If Not DoWrap Then
        BuildSignature = OneLine
        Exit Function
    End If

    ' One parameter per line while that stays within the continuation limit;
    ' otherwise pack several parameters per line, widening until it fits.
    OnePer = (PC + 1 <= MAX_PHYS_LINES)
    Limit = Opt.MaxLineLength
    If Limit < 60 Then Limit = 100

    Do
        Res = WrapParams(Head, Params, PC, Tail, PrintLevel, Opt, Limit, OnePer, LineCount)
        If LineCount <= MAX_PHYS_LINES Then Exit Do
        OnePer = False
        Limit = Limit + 20
    Loop While Limit < MAX_PHYS_CHARS

    Multi = True
    BuildSignature = Res

End Function

Private Function WrapParams(ByVal Head As String, ByRef Params() As String, ByVal PC As Long, _
                            ByVal Tail As String, ByVal PrintLevel As Long, _
                            ByRef Opt As TFormatOptions, ByVal Limit As Long, _
                            ByVal OnePer As Boolean, ByRef LineCount As Long) As String

    Dim Ind As String
    Dim Res As String
    Dim Cur As String
    Dim Piece As String
    Dim I As Long
    Dim Used As Long

    Ind = IndentString(PrintLevel + Opt.ContinuationLevels, Opt)
    Used = (PrintLevel + Opt.ContinuationLevels) * Opt.IndentSize
    Res = IndentString(PrintLevel, Opt) & Head & " _"
    LineCount = 1

    For I = 1 To PC
        Piece = Params(I)
        If I < PC Then
            Piece = Piece & ","
        Else
            Piece = Piece & ")"
            If Len(Tail) > 0 Then Piece = Piece & " " & Tail
        End If

        If OnePer Then
            Res = Res & vbCrLf & Ind & Piece
            If I < PC Then Res = Res & " _"
            LineCount = LineCount + 1
        ElseIf Len(Cur) = 0 Then
            Cur = Piece
        ElseIf Used + Len(Cur) + 1 + Len(Piece) + 2 <= Limit Then
            Cur = Cur & " " & Piece
        Else
            Res = Res & vbCrLf & Ind & Cur & " _"
            LineCount = LineCount + 1
            Cur = Piece
        End If
    Next I

    If Not OnePer Then
        Res = Res & vbCrLf & Ind & Cur
        LineCount = LineCount + 1
    End If

    WrapParams = Res

End Function

Private Function MatchParen(ByVal M As String, ByVal OpenPos As Long) As Long

    Dim I As Long
    Dim D As Long
    Dim C As String

    For I = OpenPos To Len(M)
        C = Mid$(M, I, 1)
        If C = "(" Then
            D = D + 1
        ElseIf C = ")" Then
            D = D - 1
            If D = 0 Then
                MatchParen = I
                Exit Function
            End If
        End If
    Next I

End Function

Private Function IndentString(ByVal Levels As Long, ByRef Opt As TFormatOptions) As String

    If Levels <= 0 Then Exit Function
    If Opt.UseTabs Then
        IndentString = String$(Levels, vbTab)
    Else
        IndentString = Space$(Levels * Opt.IndentSize)
    End If

End Function

' ===========================================================================
' Pass 2: blank-line layout
' ===========================================================================

Private Sub LayoutPass(ByRef A() As TOutLine, ByVal AN As Long, ByRef Opt As TFormatOptions, _
                       ByRef R() As TOutLine, ByRef RN As Long)

    Dim I As Long
    Dim Nx As Long
    Dim B As Long
    Dim Need As Long
    Dim K As Long
    Dim Blank0 As TOutLine

    Blank0.Kind = K_BLANK
    RN = 0

    I = 1
    Do While I <= AN
        If A(I).Kind = K_BLANK Then I = I + 1 Else Exit Do
    Loop

    Do While I <= AN

        PushLine R, RN, A(I)

        B = 0
        Nx = I + 1
        Do While Nx <= AN
            If A(Nx).Kind = K_BLANK Then
                B = B + 1
                Nx = Nx + 1
            Else
                Exit Do
            End If
        Loop

        If Nx > AN Then
            Need = 0
        Else
            Need = B
            If Opt.MaxBlankLines > 0 And Need > Opt.MaxBlankLines Then Need = Opt.MaxBlankLines
            If A(I).Kind = K_PROC_CLOSE And Opt.BlankLinesBetweenProcs >= 0 Then
                Need = Opt.BlankLinesBetweenProcs
            ElseIf A(I).Kind = K_DECL And Opt.BlankAfterDeclarations Then
                If Need = 0 Then
                    If NeedsBlankAfterDecl(A, AN, I, Nx) Then Need = 1
                End If
            End If
        End If

        For K = 1 To Need
            PushLine R, RN, Blank0
        Next K

        I = Nx

    Loop

End Sub

Private Function NeedsBlankAfterDecl(ByRef A() As TOutLine, ByVal AN As Long, _
                                     ByVal Cur As Long, ByVal Nx As Long) As Boolean

    Dim J As Long

    If IsOptionLine(A(Cur)) Then
        NeedsBlankAfterDecl = Not IsOptionLine(A(Nx))
        Exit Function
    End If

    If A(Nx).Kind = K_COMMENT Then
        J = Nx
        Do While J <= AN
            If A(J).Kind = K_COMMENT Then J = J + 1 Else Exit Do
        Loop
        If J <= AN Then NeedsBlankAfterDecl = IsExecutableKind(A(J).Kind)
    Else
        NeedsBlankAfterDecl = IsExecutableKind(A(Nx).Kind)
    End If

End Function

Private Function IsExecutableKind(ByVal Kind As Long) As Boolean

    Select Case Kind
        Case K_CODE, K_OPEN, K_SELECT_OPEN, K_PROC_OPEN, K_LABEL
            IsExecutableKind = True
    End Select

End Function

Private Function IsOptionLine(ByRef L As TOutLine) As Boolean

    If L.Kind = K_DECL Then IsOptionLine = (UCase$(Left$(L.Body, 7)) = "OPTION ")

End Function

Private Sub InsertOptionExplicit(ByRef A() As TOutLine, ByRef AN As Long)

    Dim I As Long
    Dim F As Long
    Dim Tmp As TOutLine
    Dim Ln As TOutLine

    For I = 1 To AN
        If A(I).Kind = K_DECL Then
            If UCase$(TrimAll(A(I).Body)) = "OPTION EXPLICIT" Then Exit Sub
        End If
    Next I

    For I = 1 To AN
        Select Case A(I).Kind
            Case K_BLANK, K_COMMENT, K_RAW
                ' keep looking
            Case Else
                F = I
                Exit For
        End Select
    Next I
    If F = 0 Then Exit Sub

    Tmp = A(AN)
    PushLine A, AN, Tmp
    For I = AN To F + 1 Step -1
        A(I) = A(I - 1)
    Next I

    Ln.Kind = K_DECL
    Ln.Level = 0
    Ln.Body = "Option Explicit"
    A(F) = Ln

End Sub

Private Sub PushLine(ByRef Arr() As TOutLine, ByRef Cnt As Long, ByRef Ln As TOutLine)

    Cnt = Cnt + 1
    If Cnt > UBound(Arr) Then ReDim Preserve Arr(1 To UBound(Arr) * 2)
    Arr(Cnt) = Ln

End Sub

' ===========================================================================
' Pass 3: alignment
' ===========================================================================

Private Sub AlignDeclarations(ByRef L() As TOutLine, ByVal N As Long)

    Dim I As Long
    Dim J As Long
    Dim K As Long
    Dim Pos As Long
    Dim MaxLeft As Long
    Dim LeftPart As String

    I = 1
    Do While I <= N
        If L(I).VarDecl Then

            J = I
            Do While J < N
                If L(J + 1).VarDecl And L(J + 1).Level = L(I).Level Then
                    J = J + 1
                Else
                    Exit Do
                End If
            Loop

            MaxLeft = 0
            For K = I To J
                Pos = AsPosition(L(K).Body)
                If Pos > 0 Then
                    LeftPart = RTrimWs(Left$(L(K).Body, Pos - 1))
                    If Len(LeftPart) > MaxLeft Then MaxLeft = Len(LeftPart)
                End If
            Next K

            For K = I To J
                Pos = AsPosition(L(K).Body)
                If Pos > 0 Then
                    LeftPart = RTrimWs(Left$(L(K).Body, Pos - 1))
                    L(K).Body = LeftPart & Space$(MaxLeft - Len(LeftPart) + 1) & Mid$(L(K).Body, Pos)
                End If
            Next K

            I = J + 1

        Else
            I = I + 1
        End If
    Loop

End Sub

Private Function AsPosition(ByVal Body As String) As Long

    AsPosition = FindWord(UCase$(MaskOnly(Body)), "AS")

End Function

Private Function MaskOnly(ByVal S As String) As String

    Dim I As Long
    Dim Ch As String
    Dim M As String
    Dim InQuote As Boolean

    M = S
    For I = 1 To Len(S)
        Ch = Mid$(S, I, 1)
        If InQuote Then
            Mid$(M, I, 1) = MASK_CHAR
            If Ch = """" Then InQuote = False
        ElseIf Ch = """" Then
            InQuote = True
            Mid$(M, I, 1) = MASK_CHAR
        End If
    Next I

    MaskOnly = M

End Function

Private Sub AlignTrailingComments(ByRef L() As TOutLine, ByVal N As Long, ByRef Opt As TFormatOptions)

    Dim I As Long
    Dim J As Long
    Dim K As Long
    Dim W As Long
    Dim MaxW As Long
    Dim Target As Long

    I = 1
    Do While I <= N
        If HasTrailingComment(L(I)) Then

            J = I
            Do While J < N
                If HasTrailingComment(L(J + 1)) Then J = J + 1 Else Exit Do
            Loop

            MaxW = 0
            For K = I To J
                W = L(K).Level * Opt.IndentSize + Len(L(K).Body)
                If W > MaxW Then MaxW = W
            Next K

            Target = MaxW + Opt.CommentGap
            If Opt.MaxLineLength <= 0 Or Target <= Opt.MaxLineLength Then
                For K = I To J
                    W = L(K).Level * Opt.IndentSize + Len(L(K).Body)
                    L(K).Gap = Space$(Target - W)
                Next K
            End If

            I = J + 1

        Else
            I = I + 1
        End If
    Loop

End Sub

Private Function HasTrailingComment(ByRef L As TOutLine) As Boolean

    Select Case L.Kind
        Case K_BLANK, K_COMMENT, K_RAW
            Exit Function
    End Select

    HasTrailingComment = (Not L.Multi) And Len(L.Comment) > 0 And Len(L.Body) > 0

End Function

' ===========================================================================
' Rendering and small helpers
' ===========================================================================

Private Function RenderLine(ByRef L As TOutLine, ByRef Opt As TFormatOptions) As String

    Dim S As String

    Select Case L.Kind
        Case K_BLANK
            RenderLine = ""
        Case K_RAW
            RenderLine = L.Body
        Case K_COMMENT
            RenderLine = IndentString(L.Level, Opt) & L.Comment
        Case Else
            If L.Multi Then
                S = L.Body
            Else
                S = IndentString(L.Level, Opt) & L.Body
            End If
            If Len(L.Comment) > 0 Then
                If Len(L.Gap) > 0 Then S = S & L.Gap Else S = S & " "
                S = S & L.Comment
            End If
            RenderLine = S
    End Select

End Function

Private Function IsWholeLineComment(ByVal T As String) As Boolean

    Dim C4 As String

    If Left$(T, 1) = "'" Then
        IsWholeLineComment = True
        Exit Function
    End If

    If UCase$(Left$(T, 3)) = "REM" Then
        C4 = Mid$(T, 4, 1)
        IsWholeLineComment = (Len(C4) = 0 Or C4 = " " Or C4 = vbTab)
    End If

End Function

Private Function IsAttributeLine(ByVal T As String) As Boolean

    IsAttributeLine = (UCase$(Left$(T, 10)) = "ATTRIBUTE ")

End Function

Private Function TrimAll(ByVal S As String) As String

    TrimAll = LTrimWs(RTrimWs(S))

End Function

Private Function RTrimWs(ByVal S As String) As String

    Dim N As Long
    Dim C As String

    N = Len(S)
    Do While N > 0
        C = Mid$(S, N, 1)
        If C = " " Or C = vbTab Or C = vbCr Or C = vbLf Then
            N = N - 1
        Else
            Exit Do
        End If
    Loop

    RTrimWs = Left$(S, N)

End Function

Private Function LTrimWs(ByVal S As String) As String

    Dim I As Long
    Dim N As Long
    Dim C As String

    N = Len(S)
    I = 1
    Do While I <= N
        C = Mid$(S, I, 1)
        If C = " " Or C = vbTab Or C = vbCr Or C = vbLf Then
            I = I + 1
        Else
            Exit Do
        End If
    Loop

    LTrimWs = Mid$(S, I)

End Function
