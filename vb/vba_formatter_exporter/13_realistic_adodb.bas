Option Explicit

Public Function RunQuery(ByVal sql As String) As Variant
    Dim cn As Object, rs As Object
    Set cn = CreateObject("ADODB.Connection")
    cn.Open "Provider=SQLOLEDB;Data Source=srv;Initial Catalog=db;Integrated Security=SSPI;"
    Set rs = cn.Execute(sql)
    If Not rs.EOF Then
        RunQuery = rs.GetRows
    Else
        RunQuery = Empty
    End If
    rs.Close: cn.Close
End Function

Public Sub BuildSql()
    Dim sql As String
    sql = "SELECT a.id, b.name " & _
          "FROM tbl a " & _
          "INNER JOIN other b ON a.id = b.id " & _
          "WHERE a.x > 1 AND b.y IN (1,2,3) " & _
          "  AND a.note = 'it''s a -- not a comment' " & _
          "ORDER BY a.id"
    Debug.Print sql
    sql = "UPDATE t SET a = 1, b = 2 WHERE id = " & CStr(5) & ";"
    sql = "DELETE FROM t WHERE name = '" & Replace("o'brien", "'", "''") & "'"
End Sub
