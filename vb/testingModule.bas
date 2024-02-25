Attribute VB_Name = "testingModule"
Option Explicit
Private Sub Test()
    Debug.Print LevenshteinDistance("kitten", "sitting")
    Debug.Print LevenshteinDistance("saturday", "sunday")
    Debug.Print LevenshteinDistance("mite", "kite")
End Sub
   

