Attribute VB_Name = "splitFunctions"
Option Explicit
'http://spreadsheetpage.com/index.php/site/tip/the_versatile_split_function/
Public Function WordCount(txt) As Long
'   Returns the number of words in a string
    Dim x As Variant
    txt = Application.Trim(txt)
    x = Split(txt, " ")
    WordCount = UBound(x) + 1
End Function

'Splitting up a filename
'The two examples in this section make it easy to extract a path or a filename from a full filespec, such as "c:\files\workbooks\archives\budget98.xls"
Public Function ExtractFileName(filespec) As String
'   Returns a filename from a filespec
    Dim x As Variant
    x = Split(filespec, Application.PathSeparator)
    ExtractFileName = x(UBound(x))
End Function

Public Function ExtractPathName(filespec) As String
'   Returns the path from a filespec
    Dim x As Variant
    x = Split(filespec, Application.PathSeparator)
    ReDim Preserve x(0 To UBound(x) - 1)
    ExtractPathName = Join(x, Application.PathSeparator) & _
      Application.PathSeparator
End Function

'Using the filespec shown above as the argument, ExtractFileName returns "budget98.xls" and ExtractPathName returns "c:\files\workbooks\archives\"
'Counting specific characters in a string
'The Public Function below accepts a string and a substring as arguments, and returns the number of times the substring is contained in the string.
Public Function CountOccurrences(str, substring) As Long
'   Returns the number of times substring appears in str
    Dim x As Variant
    x = Split(str, substring)
    CountOccurrences = UBound(x)
End Function

'Finding the longest word
'The Public Function below accepts a sentence, and returns the longest word in the sentence.
Public Function LongestWord(str) As String
' Returns the longest word in a string of words
    Dim x As Variant
    Dim i As Long
    str = Application.Trim(str)
    x = Split(str, " ")
    LongestWord = x(0)
    For i = 1 To UBound(x)
        If Len(x(i)) > Len(LongestWord) Then
            LongestWord = x(i)
        End If
    Next i
End Function

Public Function ExtractElement(str As String, delimiter As String, n As Long) As String
'   Returns the path from a filespec
    Dim x As Variant
    x = Split(str, delimiter)
    ExtractElement = x(n)
End Function

