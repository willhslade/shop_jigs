Sub Loops()
For Each c In Range("A1:A10").Cells
If c.Value = "" Then Exit For
Next c
For i = 1 To 10 Step 2
Next
For i = 1 To 3: s = s & i: Next i
With ws: .Cells(1, 1).Value = 1: End With
Do: i = i + 1: Loop While i < 5
Do Until done
done = True
Loop
Do While Not rs.EOF
rs.MoveNext
Loop
While i > 0: i = i - 1: Wend
For i = 1 To 3
For j = 1 To 3
Next j, i
With a.b(1).c
With .d
.e = 1
End With
End With
End Sub
