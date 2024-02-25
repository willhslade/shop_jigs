Option Explicit

'https://www.thespreadsheetguru.com/vba/generate-string-random-characters-vba-codet
Function RandomString(Length As Integer)
'PURPOSE: Create a Randomized String of Characters
'SOURCE: www.TheSpreadsheetGuru.com/the-code-vault

Dim CharacterBank As Variant
Dim x As Long
Dim str As String

'Test Length Input
  If Length < 1 Then
    MsgBox "Length variable must be greater than 0"
    Exit Function
  End If

CharacterBank = Array("a", "b", "c", "d", "e", "f", "g", "h", "i", "j", _
  "k", "l", "m", "n", "o", "p", "q", "r", "s", "t", "u", "v", "w", "x", _
  "y", "z", "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "!", "@", _
  "#", "$", "%", "^", "&", "*", "A", "B", "C", "D", "E", "F", "G", "H", _
  "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", _
  "W", "X", "Y", "Z")
  

'Randomly Select Characters One-by-One
  For x = 1 To Length
    Randomize
    str = str & CharacterBank(Int((UBound(CharacterBank) - LBound(CharacterBank) + 1) * Rnd + LBound(CharacterBank)))
  Next x

'Output Randomly Generated String
  RandomString = str

End Function

