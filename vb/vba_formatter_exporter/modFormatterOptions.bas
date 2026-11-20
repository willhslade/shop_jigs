Attribute VB_Name = "modFormatterOptions"
Option Explicit

' ---------------------------------------------------------------------------
' Options for the VBA formatter.
' Settings live in a plain text file so they can be edited and shared:
'     %APPDATA%\VBAFormatter\options.ini
' Run EditFormatterOptions (VBA Formatter menu) to open it in Notepad.
' ---------------------------------------------------------------------------

Public Type TFormatOptions
    IndentSize As Long              ' spaces per level (also the tab width for alignment maths)
    UseTabs As Boolean              ' indent with tabs instead of spaces
    IndentCase As Boolean           ' True: Case sits one level inside Select Case
    ContinuationLevels As Long      ' extra indent levels for " _" continuation lines
    MaxBlankLines As Long           ' collapse runs of blank lines to this many (0 = leave alone)
    BlankLinesBetweenProcs As Long  ' blank lines after End Sub/Function/Property/Type/Enum (-1 = leave alone)
    BlankAfterDeclarations As Boolean ' blank line between a declaration block and the code after it
    SignatureWrap As Long           ' 0 = leave, 1 = wrap only when too long (and unwrap when it fits), 2 = always wrap 2+ params
    MaxLineLength As Long           ' used by SignatureWrap = 1 and comment alignment (0 = no limit)
    SpaceOperators As Boolean       ' one space around = <> <= >= < >
    SpaceCommas As Boolean          ' no space before a comma, one space after
    AlignAs As Boolean              ' align "As" in consecutive Dim/Const/Private/Public lines
    AlignComments As Boolean        ' align trailing comments in consecutive lines
    CommentGap As Long              ' spaces between code and an aligned trailing comment
    AddOptionExplicit As Boolean    ' insert Option Explicit in whole-module runs when missing
    BackupBeforeBulk As Boolean     ' export modules to %TEMP% before module/project runs
End Type

Public Function DefaultOptions() As TFormatOptions

    Dim O As TFormatOptions

    O.IndentSize = 4
    O.UseTabs = False
    O.IndentCase = True
    O.ContinuationLevels = 1
    O.MaxBlankLines = 1
    O.BlankLinesBetweenProcs = 1
    O.BlankAfterDeclarations = True
    O.SignatureWrap = 1
    O.MaxLineLength = 100
    O.SpaceOperators = True
    O.SpaceCommas = True
    O.AlignAs = False
    O.AlignComments = False
    O.CommentGap = 2
    O.AddOptionExplicit = False
    O.BackupBeforeBulk = True

    DefaultOptions = O

End Function

Public Function OptionsFilePath() As String

    OptionsFilePath = Environ$("APPDATA") & "\VBAFormatter\options.ini"

End Function

Public Function LoadOptions() As TFormatOptions

    Dim O As TFormatOptions
    Dim Path As String
    Dim F As Integer
    Dim Ln As String

    O = DefaultOptions()
    Path = OptionsFilePath()

    On Error GoTo Done

    If FileExists(Path) Then
        F = FreeFile
        Open Path For Input As #F
        Do While Not EOF(F)
            Line Input #F, Ln
            ApplyOptionLine O, Ln
        Loop
    End If

Done:
    On Error Resume Next
    If F <> 0 Then Close #F
    On Error GoTo 0

    NormalizeOptions O
    LoadOptions = O

End Function

Public Sub OpenOptionsFile()

    Dim Path As String

    Path = OptionsFilePath()
    If Not FileExists(Path) Then WriteDefaultOptionsFile
    Shell "notepad.exe """ & Path & """", vbNormalFocus

End Sub

Public Sub WriteDefaultOptionsFile()

    Dim Path As String
    Dim F As Integer

    Path = OptionsFilePath()
    EnsureFolder Left$(Path, InStrRev(Path, "\") - 1)

    F = FreeFile
    Open Path For Output As #F
    Print #F, "; VBA Formatter options. Lines starting with ; are comments."
    Print #F, "; Booleans: true/false. Changes apply the next time you run the formatter."
    Print #F, ""
    Print #F, "; Spaces per indent level (also used as tab width for alignment)."
    Print #F, "IndentSize=4"
    Print #F, "; Indent with tabs instead of spaces."
    Print #F, "UseTabs=false"
    Print #F, "; true: Case is indented one level inside Select Case (VBE default look)."
    Print #F, "IndentCase=true"
    Print #F, "; Extra indent levels for continuation lines (lines ending in "" _"")."
    Print #F, "ContinuationLevels=1"
    Print #F, "; Collapse runs of blank lines to this many. 0 = leave as is."
    Print #F, "MaxBlankLines=1"
    Print #F, "; Blank lines after End Sub / End Function / End Property / End Type / End Enum. -1 = leave as is."
    Print #F, "BlankLinesBetweenProcs=1"
    Print #F, "; Blank line between a block of Dim/Const/Option lines and the code that follows."
    Print #F, "BlankAfterDeclarations=true"
    Print #F, "; Procedure signatures: 0 = leave, 1 = wrap only when longer than MaxLineLength (unwrap when it fits), 2 = always wrap when 2+ parameters."
    Print #F, "SignatureWrap=1"
    Print #F, "; Maximum line length for signature wrapping and comment alignment. 0 = no limit."
    Print #F, "MaxLineLength=100"
    Print #F, "; One space around = <> <= >= < > (strings and comments are never touched)."
    Print #F, "SpaceOperators=true"
    Print #F, "; No space before a comma, one space after."
    Print #F, "SpaceCommas=true"
    Print #F, "; Align the word As in consecutive Dim/Const/Private/Public lines."
    Print #F, "AlignAs=false"
    Print #F, "; Align trailing comments in consecutive lines."
    Print #F, "AlignComments=false"
    Print #F, "; Spaces between code and an aligned trailing comment."
    Print #F, "CommentGap=2"
    Print #F, "; Insert Option Explicit when formatting a whole module and it is missing."
    Print #F, "AddOptionExplicit=false"
    Print #F, "; Export modules to %TEMP%\VBAFormatterBackup before module/project runs."
    Print #F, "BackupBeforeBulk=true"
    Close #F

End Sub

Public Sub EnsureFolder(ByVal Path As String)

    Dim Fso As Object

    If Len(Path) = 0 Then Exit Sub
    Set Fso = CreateObject("Scripting.FileSystemObject")
    If Fso.FolderExists(Path) Then Exit Sub

    EnsureFolder Fso.GetParentFolderName(Path)
    Fso.CreateFolder Path

End Sub

Private Function FileExists(ByVal Path As String) As Boolean

    On Error Resume Next
    FileExists = (Len(Dir$(Path)) > 0)

End Function

Private Sub ApplyOptionLine(ByRef O As TFormatOptions, ByVal Ln As String)

    Dim P As Long
    Dim K As String
    Dim V As String

    Ln = Trim$(Ln)
    If Len(Ln) = 0 Then Exit Sub
    If Left$(Ln, 1) = ";" Or Left$(Ln, 1) = "#" Then Exit Sub

    P = InStr(Ln, "=")
    If P < 2 Then Exit Sub

    K = LCase$(Trim$(Left$(Ln, P - 1)))
    V = Trim$(Mid$(Ln, P + 1))

    P = InStr(V, ";")
    If P > 0 Then V = Trim$(Left$(V, P - 1))

    Select Case K
        Case "indentsize":              O.IndentSize = CLng(Val(V))
        Case "usetabs":                 O.UseTabs = ToBool(V)
        Case "indentcase":              O.IndentCase = ToBool(V)
        Case "continuationlevels":      O.ContinuationLevels = CLng(Val(V))
        Case "maxblanklines":           O.MaxBlankLines = CLng(Val(V))
        Case "blanklinesbetweenprocs":  O.BlankLinesBetweenProcs = CLng(Val(V))
        Case "blankafterdeclarations":  O.BlankAfterDeclarations = ToBool(V)
        Case "signaturewrap":           O.SignatureWrap = CLng(Val(V))
        Case "maxlinelength":           O.MaxLineLength = CLng(Val(V))
        Case "spaceoperators":          O.SpaceOperators = ToBool(V)
        Case "spacecommas":             O.SpaceCommas = ToBool(V)
        Case "alignas":                 O.AlignAs = ToBool(V)
        Case "aligncomments":           O.AlignComments = ToBool(V)
        Case "commentgap":              O.CommentGap = CLng(Val(V))
        Case "addoptionexplicit":       O.AddOptionExplicit = ToBool(V)
        Case "backupbeforebulk":        O.BackupBeforeBulk = ToBool(V)
    End Select

End Sub

Private Function ToBool(ByVal V As String) As Boolean

    Select Case LCase$(Trim$(V))
        Case "1", "true", "yes", "on", "y"
            ToBool = True
    End Select

End Function

Private Sub NormalizeOptions(ByRef O As TFormatOptions)

    If O.IndentSize < 1 Then O.IndentSize = 1
    If O.IndentSize > 16 Then O.IndentSize = 16
    If O.ContinuationLevels < 0 Then O.ContinuationLevels = 0
    If O.ContinuationLevels > 4 Then O.ContinuationLevels = 4
    If O.MaxBlankLines < 0 Then O.MaxBlankLines = 0
    If O.BlankLinesBetweenProcs < -1 Then O.BlankLinesBetweenProcs = -1
    If O.BlankLinesBetweenProcs > 5 Then O.BlankLinesBetweenProcs = 5
    If O.SignatureWrap < 0 Or O.SignatureWrap > 2 Then O.SignatureWrap = 1
    If O.MaxLineLength < 0 Then O.MaxLineLength = 0
    If O.MaxLineLength > 0 And O.MaxLineLength < 40 Then O.MaxLineLength = 40
    If O.CommentGap < 1 Then O.CommentGap = 1

End Sub
