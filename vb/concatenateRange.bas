Attribute VB_Name = "concatenateRange"
'http://www.vbaexpress.com/kb/getarticle.php?kb_id=817
Option Explicit

Public Function ConcRange(Substrings As Range, Optional Delim As String = "", _
    Optional AsDisplayed As Boolean = False, Optional SkipBlanks As Boolean = False)
     
     ' Concatenates a range of cells, using an optional delimiter.  The concatenated
     ' strings may be either actual values (AsDisplayed=False) or displayed values.
     ' If NoBlanks=True, blanks cells or cells that evaluate to a zero-length string
     ' are skipped in the concatenation
     
     ' Substrings: the range of cells whose values/text you want to concatenate.  May be
     ' from a row, a column, or a "rectangular" range (1+ rows, 1+ columns)
     
     ' Delimiter: the optional separator you want inserted between each item to be
     ' concatenated.  By default, the function will use a zero-length string as the
     ' delimiter (which is what Excel's CONCATENATE function does), but you can specify
     ' your own character(s).  (The Delimiter can be more than one character)
     
     ' AsDisplayed: for numeric values (includes currency but not dates), this controls
     ' whether the real value of the cell is used for concatenation, or the formatted
     ' displayed value.  Note for how dates are handled: if AsDisplayed is FALSE or omitted,
     ' dates will show up using whatever format you have selected in your regional settings
     ' for displaying dates.  If AsDisplayed=TRUE, dates will use the formatted displayed
     ' value
     
     ' SkipBlanks: Indicates whether the function should ignore blank cells (or cells with
     ' nothing but spaces) in the Substrings range when it performs the concatenation.
     ' If NoBlanks=FALSE or is omitted, the function includes blank cells in the
     ' concatenation.  In the examples above, where NoBlanks=False, you will see "extra"
     ' delimiters in cases where the Substrings range has blank cells (or cells with only
     ' spaces)
     
    Dim CLL As Range
     
    For Each CLL In Substrings.Cells
        If Not (SkipBlanks And Trim(CLL) = "") Then
            ConcRange = ConcRange & Delim & IIf(AsDisplayed, Trim(CLL.Text), Trim(CLL.Value))
        End If
    Next CLL
     
    ConcRange = Mid$(ConcRange, Len(Delim) + 1)
     
End Function

