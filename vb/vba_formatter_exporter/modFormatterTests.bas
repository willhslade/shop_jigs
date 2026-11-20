Attribute VB_Name = "modFormatterTests"
Option Explicit

' ---------------------------------------------------------------------------
' Self-tests. Run RunFormatterTests, results go to the Immediate window (Ctrl+G).
' Every case checks (a) output equals the expected text and (b) formatting the
' output again changes nothing (idempotence).
' ---------------------------------------------------------------------------

Private Pass As Long
Private Fail As Long

Public Sub RunFormatterTests()

    Dim Opt As TFormatOptions

    Pass = 0
    Fail = 0

    Opt = DefaultOptions()
    Opt.AddOptionExplicit = False

    TestBasicIndent Opt
    TestEndFunctionDoesNotDrift Opt
    TestIdentifiersNotKeywords Opt
    TestStringsAndComments Opt
    TestSelectCase Opt
    TestOneLineIf Opt
    TestElseIfChain Opt
    TestContinuation Opt
    TestSignatureWrapAndUnwrap Opt
    TestBlankLines Opt
    TestDirectivesAndLabels Opt
    TestPropertiesAndModifiers Opt
    TestDeclare Opt
    TestSpacing Opt
    TestTabs Opt
    TestAlignment Opt
    TestOptionExplicitInsert Opt
    TestFragment Opt
    TestInlineBlocks Opt
    TestContinuationLimit Opt
    TestCommentContinuation Opt
    TestNumericLabel Opt

    Debug.Print String$(50, "-")
    Debug.Print "Formatter tests: " & Pass & " passed, " & Fail & " failed."
    MsgBox "Formatter tests: " & Pass & " passed, " & Fail & " failed." & vbCrLf & _
           "Details in the Immediate window.", IIf(Fail = 0, vbInformation, vbExclamation), "VBA Formatter"

End Sub

' ---------------------------------------------------------------------------

Private Function L(ParamArray Parts() As Variant) As String

    Dim I As Long
    Dim S As String

    For I = LBound(Parts) To UBound(Parts)
        If I > LBound(Parts) Then S = S & vbCrLf
        S = S & Parts(I)
    Next I

    L = S

End Function

Private Sub Check(ByVal CaseName As String, ByVal InText As String, ByVal Expected As String, _
                  ByRef Opt As TFormatOptions)

    Dim Out1 As String
    Dim Out2 As String

    Out1 = FormatWithOptions(InText, Opt)
    Out2 = FormatWithOptions(Out1, Opt)

    If Out1 = Expected Then
        Pass = Pass + 1
    Else
        Fail = Fail + 1
        Debug.Print "FAIL  " & CaseName
        Debug.Print "  expected:" & vbCrLf & Visible(Expected)
        Debug.Print "  actual:" & vbCrLf & Visible(Out1)
    End If

    If Out2 = Out1 Then
        Pass = Pass + 1
    Else
        Fail = Fail + 1
        Debug.Print "FAIL  " & CaseName & "  (not idempotent)"
        Debug.Print "  first:" & vbCrLf & Visible(Out1)
        Debug.Print "  second:" & vbCrLf & Visible(Out2)
    End If

End Sub

Private Function Visible(ByVal S As String) As String

    Visible = "    |" & Replace(Replace(S, vbTab, "<TAB>"), vbCrLf, "|" & vbCrLf & "    |") & "|"

End Function

' ---------------------------------------------------------------------------

Private Sub TestBasicIndent(ByRef Opt As TFormatOptions)

    Check "basic indent", _
        L("Sub A()", "Dim x As Long", "For x = 1 To 3", "Debug.Print x", "Next x", "End Sub"), _
        L("Sub A()", "    Dim x As Long", "", "    For x = 1 To 3", "        Debug.Print x", "    Next x", "End Sub"), Opt

End Sub

Private Sub TestEndFunctionDoesNotDrift(ByRef Opt As TFormatOptions)

    Check "End Function / Exit Function", _
        L("Function F() As Long", "    If True Then Exit Function", "    F = 1", "End Function", _
          "", "Function G() As Long", "G = 2", "End Function"), _
        L("Function F() As Long", "    If True Then Exit Function", "    F = 1", "End Function", _
          "", "Function G() As Long", "    G = 2", "End Function"), Opt

End Sub

Private Sub TestIdentifiersNotKeywords(ByRef Opt As TFormatOptions)

    Check "identifiers that start like keywords", _
        L("Sub A()", "x = IfError(y, 0)", "DoEvents", "SubTotal = 1", "ForeColor = 2", _
          "WithEvents_ = 3", "NextValue = 4", "End Sub"), _
        L("Sub A()", "    x = IfError(y, 0)", "    DoEvents", "    SubTotal = 1", "    ForeColor = 2", _
          "    WithEvents_ = 3", "    NextValue = 4", "End Sub"), Opt

End Sub

Private Sub TestStringsAndComments(ByRef Opt As TFormatOptions)

    Check "keywords in strings and comments", _
        L("Sub A()", "s = ""If x Then"" ' End Sub", "' If this is a comment", "Rem For each", _
          "t = ""it""""s Next""", "End Sub"), _
        L("Sub A()", "    s = ""If x Then"" ' End Sub", "    ' If this is a comment", "    Rem For each", _
          "    t = ""it""""s Next""", "End Sub"), Opt

End Sub

Private Sub TestSelectCase(ByRef Opt As TFormatOptions)

    Dim O2 As TFormatOptions

    Check "select case (indented Case)", _
        L("Sub A()", "Select Case x", "Case 1", "y = 1", "Case Else", "y = 2", "End Select", "End Sub"), _
        L("Sub A()", "    Select Case x", "        Case 1", "            y = 1", "        Case Else", _
          "            y = 2", "    End Select", "End Sub"), Opt

    O2 = Opt
    O2.IndentCase = False
    Check "select case (flat Case)", _
        L("Sub A()", "Select Case x", "Case 1", "y = 1", "End Select", "End Sub"), _
        L("Sub A()", "    Select Case x", "    Case 1", "        y = 1", "    End Select", "End Sub"), O2

End Sub

Private Sub TestOneLineIf(ByRef Opt As TFormatOptions)

    Check "one-line If", _
        L("Sub A()", "If x Then y = 1", "If x Then y = 1 Else y = 2", "z = 3", "End Sub"), _
        L("Sub A()", "    If x Then y = 1", "    If x Then y = 1 Else y = 2", "    z = 3", "End Sub"), Opt

    Check "If with continuation is a block", _
        L("Sub A()", "If a And _", "b Then", "x = 1", "End If", "End Sub"), _
        L("Sub A()", "    If a And _", "        b Then", "        x = 1", "    End If", "End Sub"), Opt

End Sub

Private Sub TestElseIfChain(ByRef Opt As TFormatOptions)

    Check "ElseIf / Else", _
        L("Sub A()", "If a Then", "x = 1", "ElseIf b Then", "x = 2", "Else", "x = 3", "End If", "End Sub"), _
        L("Sub A()", "    If a Then", "        x = 1", "    ElseIf b Then", "        x = 2", "    Else", _
          "        x = 3", "    End If", "End Sub"), Opt

End Sub

Private Sub TestContinuation(ByRef Opt As TFormatOptions)

    Check "continuation lines", _
        L("Sub A()", "x = a + _", "b + _", "c", "y = 1", "End Sub"), _
        L("Sub A()", "    x = a + _", "        b + _", "        c", "    y = 1", "End Sub"), Opt

End Sub

Private Sub TestSignatureWrapAndUnwrap(ByRef Opt As TFormatOptions)

    Dim O2 As TFormatOptions

    Check "short signature stays on one line", _
        L("Public Sub A(ByVal x As Long, ByVal y As Long)", "End Sub"), _
        L("Public Sub A(ByVal x As Long, ByVal y As Long)", "End Sub"), Opt

    Check "wrapped signature that fits is joined", _
        L("Public Sub A( _", "    ByVal x As Long, _", "    ByVal y As Long)", "End Sub"), _
        L("Public Sub A(ByVal x As Long, ByVal y As Long)", "End Sub"), Opt

    O2 = Opt
    O2.MaxLineLength = 60
    Check "long signature wraps with continuations", _
        L("Public Function DoSomethingLong(ByVal FirstArgument As String, ByVal Second As Long, Optional ByVal Third As String = ""a,b"") As Boolean", _
          "End Function"), _
        L("Public Function DoSomethingLong( _", "    ByVal FirstArgument As String, _", "    ByVal Second As Long, _", _
          "    Optional ByVal Third As String = ""a,b"") As Boolean", "End Function"), O2

    O2.SignatureWrap = 2
    Check "always-wrap mode", _
        L("Sub A(ByVal x As Long, ByVal y As Long)", "End Sub"), _
        L("Sub A( _", "    ByVal x As Long, _", "    ByVal y As Long)", "End Sub"), O2

End Sub

Private Sub TestBlankLines(ByRef Opt As TFormatOptions)

    Check "blank lines collapse and proc spacing", _
        L("Sub A()", "x = 1", "", "", "", "y = 2", "End Sub", "Sub B()", "End Sub", "", "", "", "", ""), _
        L("Sub A()", "    x = 1", "", "    y = 2", "End Sub", "", "Sub B()", "End Sub"), Opt

End Sub

Private Sub TestDirectivesAndLabels(ByRef Opt As TFormatOptions)

    Check "directives and labels flush left", _
        L("Sub A()", "On Error GoTo EH", "x = 1", "#If VBA7 Then", "y = 2", "#End If", "Exit Sub", "EH:", _
          "MsgBox ""x""", "End Sub"), _
        L("Sub A()", "    On Error GoTo EH", "    x = 1", "#If VBA7 Then", "    y = 2", "#End If", _
          "    Exit Sub", "EH:", "    MsgBox ""x""", "End Sub"), Opt

End Sub

Private Sub TestPropertiesAndModifiers(ByRef Opt As TFormatOptions)

    Check "properties, Friend, Static, Enum, Type", _
        L("Property Get P() As Long", "P = 1", "End Property", "Friend Sub F()", "End Sub", _
          "Static Function S() As Long", "S = 1", "End Function", _
          "Public Enum E", "A = 1", "End Enum", "Private Type T", "X As Long", "End Type"), _
        L("Property Get P() As Long", "    P = 1", "End Property", "", "Friend Sub F()", "End Sub", "", _
          "Static Function S() As Long", "    S = 1", "End Function", "", _
          "Public Enum E", "    A = 1", "End Enum", "", "Private Type T", "    X As Long", "End Type"), Opt

End Sub

Private Sub TestDeclare(ByRef Opt As TFormatOptions)

    Check "Declare is not a block", _
        L("Private Declare Function GetTickCount Lib ""kernel32"" () As Long", "Sub A()", "x = 1", "End Sub"), _
        L("Private Declare Function GetTickCount Lib ""kernel32"" () As Long", "", "Sub A()", "    x = 1", "End Sub"), Opt

End Sub

Private Sub TestSpacing(ByRef Opt As TFormatOptions)

    Check "operator and comma spacing", _
        L("Sub A()", "If a<>b And c>=d Then", "x=Foo(1,2 ,3)", "y = ""a=b,c""", "Call Bar(Name:=1,Other:=2)", _
          "End If", "End Sub"), _
        L("Sub A()", "    If a <> b And c >= d Then", "        x = Foo(1, 2, 3)", "        y = ""a=b,c""", _
          "        Call Bar(Name:=1, Other:=2)", "    End If", "End Sub"), Opt

End Sub

Private Sub TestTabs(ByRef Opt As TFormatOptions)

    Dim O2 As TFormatOptions

    O2 = Opt
    O2.UseTabs = True
    Check "tab indentation", _
        L("Sub A()", "If x Then", "y = 1", "End If", "End Sub"), _
        L("Sub A()", vbTab & "If x Then", vbTab & vbTab & "y = 1", vbTab & "End If", "End Sub"), O2

End Sub

Private Sub TestAlignment(ByRef Opt As TFormatOptions)

    Dim O2 As TFormatOptions

    O2 = Opt
    O2.AlignAs = True
    O2.AlignComments = True
    Check "align As and comments", _
        L("Sub A()", "Dim a As Long ' one", "Dim longName As String ' two", "x = 1", "End Sub"), _
        L("Sub A()", "    Dim a        As Long    ' one", "    Dim longName As String  ' two", "", "    x = 1", "End Sub"), O2

End Sub

Private Sub TestOptionExplicitInsert(ByRef Opt As TFormatOptions)

    Dim O2 As TFormatOptions

    O2 = Opt
    O2.AddOptionExplicit = True
    Check "insert Option Explicit", _
        L("Sub A()", "End Sub"), _
        L("Option Explicit", "", "Sub A()", "End Sub"), O2

End Sub

Private Sub TestFragment(ByRef Opt As TFormatOptions)

    Dim Got As String

    Got = FormatWithOptions(L("x = 1", "If a Then", "y = 2", "End If"), Opt, 1, True)
    If Got = L("    x = 1", "    If a Then", "        y = 2", "    End If") Then
        Pass = Pass + 1
    Else
        Fail = Fail + 1
        Debug.Print "FAIL  fragment at level 1" & vbCrLf & Visible(Got)
    End If

    Got = FormatWithOptions(L("ElseIf b Then", "z = 3"), Opt, 1, True)
    If Got = L("    ElseIf b Then", "        z = 3") Then
        Pass = Pass + 1
    Else
        Fail = Fail + 1
        Debug.Print "FAIL  fragment starting with ElseIf" & vbCrLf & Visible(Got)
    End If

End Sub

Private Sub TestInlineBlocks(ByRef Opt As TFormatOptions)

    Check "inline For..Next and With..End With", _
        L("Sub A()", "For i = 1 To 3: x = x + i: Next i", "With y: .a = 1: End With", "z = 1", "End Sub"), _
        L("Sub A()", "    For i = 1 To 3: x = x + i: Next i", "    With y: .a = 1: End With", "    z = 1", "End Sub"), Opt

End Sub

Private Sub TestContinuationLimit(ByRef Opt As TFormatOptions)

    ' 30 parameters must not produce more than 24 continuations (25 physical lines)
    Dim O2 As TFormatOptions
    Dim Src As String
    Dim Out1 As String
    Dim Out2 As String
    Dim Parts() As String
    Dim I As Long
    Dim N As Long

    Src = "Sub Many("
    For I = 1 To 30
        If I > 1 Then Src = Src & ", "
        Src = Src & "ByVal Parameter" & I & " As Long"
    Next I
    Src = Src & ")" & vbCrLf & "End Sub"

    O2 = Opt
    O2.SignatureWrap = 2

    Out1 = FormatWithOptions(Src, O2)
    Out2 = FormatWithOptions(Out1, O2)
    Parts = Split(Out1, vbCrLf)

    N = 0
    Do While N <= UBound(Parts)
        If Trim$(Parts(N)) = "End Sub" Then Exit Do
        N = N + 1
    Loop

    If N <= 25 And N >= 2 Then
        Pass = Pass + 1
    Else
        Fail = Fail + 1
        Debug.Print "FAIL  signature uses " & N & " physical lines (max 25)" & vbCrLf & Visible(Out1)
    End If

    If Out1 = Out2 Then
        Pass = Pass + 1
    Else
        Fail = Fail + 1
        Debug.Print "FAIL  many-parameter signature not idempotent"
    End If

End Sub

Private Sub TestCommentContinuation(ByRef Opt As TFormatOptions)

    ' a comment ending in " _" swallows the next line, so it must not be indented as code
    Check "comment continuation", _
        L("Sub A()", "' note _", "If x Then", "y = 1", "End Sub"), _
        L("Sub A()", "    ' note _", "    If x Then", "    y = 1", "End Sub"), Opt

End Sub

Private Sub TestNumericLabel(ByRef Opt As TFormatOptions)

    Check "numeric line label stays at column 0", _
        L("Sub A()", "10 x = 1", "Dim q As Long", "End Sub"), _
        L("Sub A()", "10 x = 1", "    Dim q As Long", "End Sub"), Opt

End Sub
