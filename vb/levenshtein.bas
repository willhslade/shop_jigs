Attribute VB_Name = "levenshtein"
Option Explicit

Public Function levenshtein1(ByVal string1 As String, ByVal string2 As String) As Long
'http://stackoverflow.com/questions/4243036/levenshtein-distance-in-excel
    Dim i As Long, j As Long
    Dim string1_length As Long
    Dim string2_length As Long
    Dim distance() As Long
    
    string1_length = Len(string1)
    string2_length = Len(string2)
    ReDim distance(string1_length, string2_length)
    
    For i = 0 To string1_length
        distance(i, 0) = i
    Next
    
    For j = 0 To string2_length
        distance(0, j) = j
    Next
    
    For i = 1 To string1_length
        For j = 1 To string2_length
            If Asc(Mid$(string1, i, 1)) = Asc(Mid$(string2, j, 1)) Then
                distance(i, j) = distance(i - 1, j - 1)
            Else
                distance(i, j) = Application.WorksheetFunction.Min _
                    (distance(i - 1, j) + 1, _
                    distance(i, j - 1) + 1, _
                    distance(i - 1, j - 1) + 1)
            End If
        Next
    Next
    
    levenshtein1 = distance(string1_length, string2_length)
End Function

Public Function FuzzyMatch(ByVal string1 As String, _
    ByVal string2 As String, _
    Optional min_percentage As Long = 70) As String
'http://stackoverflow.com/questions/4243036/levenshtein-distance-in-excel
    Dim i As Long, j As Long
    Dim string1_length As Long
    Dim string2_length As Long
    Dim distance() As Long, result As Long
    
    string1_length = Len(string1)
    string2_length = Len(string2)
    
    ' Check if not too long
    If string1_length >= string2_length * (min_percentage / 100) Then
    ' Check if not too short
        If string1_length <= string2_length * ((200 - min_percentage) / 100) Then
            ReDim distance(string1_length, string2_length)
            For i = 0 To string1_length: distance(i, 0) = i: Next
            For j = 0 To string2_length: distance(0, j) = j: Next
            For i = 1 To string1_length
                For j = 1 To string2_length
                    If Asc(Mid$(string1, i, 1)) = Asc(Mid$(string2, j, 1)) Then
                        distance(i, j) = distance(i - 1, j - 1)
                    Else
                        distance(i, j) = Application.WorksheetFunction.Min _
                            (distance(i - 1, j) + 1, _
                            distance(i, j - 1) + 1, _
                            distance(i - 1, j - 1) + 1)
                    End If
                Next
            Next
            result = distance(string1_length, string2_length)
        'The distance
        End If
    End If
    
    If result <> 0 Then
        FuzzyMatch = (CLng((100 - ((result / string1_length) * 100)))) & _
            "% (" & result & ")" 'Convert to percentage
    Else
        FuzzyMatch = "Not a match"
    End If
End Function

Private Function levenshtein2(a As String, b As String) As Integer
'http://en.wikibooks.org/wiki/Algorithm_Implementation/Strings/Levenshtein_distance#Visual_Basic_for_Applications_.28no_Damerau_extension.29
    Dim i As Integer
    Dim j As Integer
    Dim cost As Integer
    Dim d() As Integer
    Dim min1 As Integer
    Dim min2 As Integer
    Dim min3 As Integer

    If Len(a) = 0 Then
        levenshtein2 = Len(b)
        Exit Function
    End If

    If Len(b) = 0 Then
        levenshtein2 = Len(a)
        Exit Function
    End If

    ReDim d(Len(a), Len(b))

    For i = 0 To Len(a)
        d(i, 0) = i
    Next

    For j = 0 To Len(b)
        d(0, j) = j
    Next

    For i = 1 To Len(a)
        For j = 1 To Len(b)
            If Mid(a, i, 1) = Mid(b, j, 1) Then
                cost = 0
            Else
                cost = 1
            End If

            ' Since Min() function is not a part of VBA, we'll "emulate" it below
            min1 = (d(i - 1, j) + 1)
            min2 = (d(i, j - 1) + 1)
            min3 = (d(i - 1, j - 1) + cost)

'            If min1 <= min2 And min1 <= min3 Then
'                d(i, j) = min1
'            ElseIf min2 <= min1 And min2 <= min3 Then
'                d(i, j) = min2
'            Else
'                d(i, j) = min3
'            End If
            
            ' In Excel we can use Min() function that is included
            ' as a method of WorksheetFunction object
            d(i, j) = Application.WorksheetFunction.Min(min1, min2, min3)
        Next
    Next

    levenshtein2 = d(Len(a), Len(b))
End Function
   
Public Function LevenshteinDistance(word1, word2)
'http://wiki.lessthandot.com/index.php/Comparing_Words:_Levenshtein_Distance
Dim s As Variant
Dim t As Variant
Dim d As Variant
Dim m, n
Dim i, j, k
Dim a(2), r
Dim cost
    m = Len(word1)
    n = Len(word2)
    
    'This is the only way to use
    'variables to dimension an array
    ReDim s(m)
    ReDim t(n)
    ReDim d(m, n)
    
    For i = 1 To m
        s(i) = Mid(word1, i, 1)
    Next
    
    For i = 1 To n
        t(i) = Mid(word2, i, 1)
    Next
    
    For i = 0 To m
        d(i, 0) = i
    Next
    
    For j = 0 To n
        d(0, j) = j
    Next
        
     
    For i = 1 To m
        For j = 1 To n
                  
            If s(i) = t(j) Then
                cost = 0
            Else
                cost = 1
            End If
            
            a(0) = d(i - 1, j) + 1             '  // deletion
            a(1) = d(i, j - 1) + 1             '  // insertion
            a(2) = d(i - 1, j - 1) + cost      '  // substitution
            
            r = a(0)
            
            For k = 1 To UBound(a)
                If a(k) < r Then r = a(k)
            Next
            
            d(i, j) = r
        
        Next
    
    Next
     
    LevenshteinDistance = d(m, n)
End Function

Private Function levenshtein3(a As String, b As String) As Integer
'http://www.access-programmers.co.uk/forums/showthread.php?t=190907
Dim i As Integer
Dim j As Integer
Dim cost As Integer
Dim d() As Integer
Dim min1 As Integer
Dim min2 As Integer
Dim min3 As Integer
    
    
    If Len(a) = 0 Then
        levenshtein3 = Len(b)
        Exit Function
    End If
    
    
    If Len(b) = 0 Then
        levenshtein3 = Len(a)
        Exit Function
    End If
    
    ReDim d(Len(a), Len(b))
    
    For i = 0 To Len(a)
        d(i, 0) = i
    Next
    
    For j = 0 To Len(b)
        d(0, j) = j
    Next
    
    
    For i = 1 To Len(a)
    
        For j = 1 To Len(b)
        
        If Mid(a, i, 1) = Mid(b, j, 1) Then
            cost = 0
        Else
            cost = 1
        End If
        
        ' Since Min() function is not a part of
        ' VBA, we'll "emulate" it below
        min1 = (d(i - 1, j) + 1)
        min2 = (d(i, j - 1) + 1)
        min3 = (d(i - 1, j - 1) + cost)
        
        
        If min1 <= min2 And min1 <= min3 Then
            d(i, j) = min1
        ElseIf min2 <= min1 And min2 <= min3 Then
            d(i, j) = min2
        Else
            d(i, j) = min3
        End If
        
        ' In Excel we can use Min() function tha
        ' t is included
        ' as a method of WorksheetFunction objec
        ' t
        'd(i, j) = Application.WorksheetFunction
        ' .Min(min1, min2, min3)
        Next
    Next
    
    levenshtein3 = d(Len(a), Len(b))
    
End Function


