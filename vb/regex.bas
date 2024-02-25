Attribute VB_Name = "regex"
Option Explicit
 
'http://vbaexpress.com/kb/getarticle.php?kb_id=841
Function RegExpFind(LookIn As String, PatternStr As String, Optional Pos, _
    Optional MatchCase As Boolean = True)
     
     ' This function uses Regular Expressions to parse a string (LookIn), and return matches to a
     ' pattern (PatternStr).  Use Pos to indicate which match you want:
     ' Pos omitted               : function returns a zero-based array of all matches
     ' Pos = 0                   : the last match
     ' Pos = 1                   : the first match
     ' Pos = 2                   : the second match
     ' Pos = <positive integer>  : the Nth match
     ' If Pos is greater than the number of matches, is negative, or is non-numeric, the function
     ' returns an empty string.  If no match is found, the function returns an empty string
     
     ' If MatchCase is omitted or True (default for RegExp) then the Pattern must match case (and
     ' thus you may have to use [a-zA-Z] instead of just [a-z] or [A-Z]).
     
     ' If you use this function in Excel, you can use range references for any of the arguments.
     ' If you use this in Excel and return the full array, make sure to set up the formula as an
     ' array formula.  If you need the array formula to go down a column, use TRANSPOSE()
     
    Dim RegX As Object
    Dim TheMatches As Object
    Dim Answer() As String
    Dim Counter As Long
     
     ' Evaluate Pos.  If it is there, it must be numeric and converted to Long
    If Not IsMissing(Pos) Then
        If Not IsNumeric(Pos) Then
            RegExpFind = ""
            Exit Function
        Else
            Pos = CLng(Pos)
        End If
    End If
     
     ' Create instance of RegExp object
    Set RegX = CreateObject("VBScript.RegExp")
    With RegX
        .Pattern = PatternStr
        .Global = True
        .IgnoreCase = Not MatchCase
    End With
     
     ' Test to see if there are any matches
    If RegX.Test(LookIn) Then
         
         ' Run RegExp to get the matches, which are returned as a zero-based collection
        Set TheMatches = RegX.Execute(LookIn)
         
         ' If Pos is missing, user wants array of all matches.  Build it and assign it as the
         ' function's return value
        If IsMissing(Pos) Then
            ReDim Answer(0 To TheMatches.Count - 1) As String
            For Counter = 0 To UBound(Answer)
                Answer(Counter) = TheMatches(Counter)
            Next
            RegExpFind = Answer
             
             ' User wanted the Nth match (or last match, if Pos = 0).  Get the Nth value, if possible
        Else
            Select Case Pos
            Case 0 ' Last match
                RegExpFind = TheMatches(TheMatches.Count - 1)
            Case 1 To TheMatches.Count ' Nth match
                RegExpFind = TheMatches(Pos - 1)
            Case Else ' Invalid item number
                RegExpFind = ""
            End Select
        End If
         
         ' If there are no matches, return empty string
    Else
        RegExpFind = ""
    End If
     
     ' Release object variables
    Set RegX = Nothing
    Set TheMatches = Nothing
     
End Function
 
Function RegExpReplace(LookIn As String, PatternStr As String, Optional ReplaceWith As String = "", _
    Optional ReplaceAll As Boolean = True, Optional MatchCase As Boolean = True)
     
     ' This function uses Regular Expressions to parse a string, and replace parts of the string
     ' matching the specified pattern with another string.  The optional argument ReplaceAll controls
     ' whether all instances of the matched string are replaced (True) or just the first instance (False)
     
     ' By default, RegExp is case-sensitive in pattern-matching.  To keep this, omit MatchCase or
     ' set it to True
     
     ' If you use this function from Excel, you may substitute range references for all the arguments
     
    Dim RegX As Object
     
    Set RegX = CreateObject("VBScript.RegExp")
    With RegX
        .Pattern = PatternStr
        .Global = ReplaceAll
        .IgnoreCase = Not MatchCase
    End With
     
    RegExpReplace = RegX.Replace(LookIn, ReplaceWith)
     
    Set RegX = Nothing
     
End Function
