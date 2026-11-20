Sub KeywordLikeNames()
    Dim IfCount As Long, DoCount As Long, ForCount As Long, EndTotal As Long
    Dim SubTotal As Long, FunctionName As String, WithValue As Long, NextRow As Long
    Dim ElseValue As Long, CaseId As Long, SelectedItem As Long, LoopCount As Long
    IfCount = 1
    DoCount = 2
    ForCount = 3
    EndTotal = 4
    SubTotal = 5
    NextRow = IfCount + DoCount
    WithValue = 1
    ElseValue = 2
    CaseId = 3
    LoopCount = 4
    FunctionName = "Function"
    DoEvents
    Application.WorksheetFunction.IfError 1, 2
    Set rng = Range("A1")
    rng.Offset(1, 0).Value = "x"
    IfCount = IIf(IfCount > 0, 1, 0)
    ThisWorkbook.Sheets("If").Range("A1").Value = "For"
    Dim PropertyName As String
    PropertyName = "Property"
    Dim TypeName2 As String
    TypeName2 = TypeName(IfCount)
    Dim EnumValue As Long, DeclareIt As Long
    EnumValue = 1: DeclareIt = 2
End Sub

Public Property Get EndDate() As Date
    EndDate = Date
End Property

Public Property Get IfValue() As Long
    IfValue = 1
End Property
