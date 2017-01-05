Sub fetchSQLData()

Const sqlconstr As String = "Provider=SQLOLEDB;User ID=USERID;Password=PASSWORD;Initial Catalog=TABLE ENTRY;Data Source=ADDRESS"
Dim cnn As New ADODB.connection
Dim rst As New ADODB.Recordset

Dim StrQuery As String

cnn.Open (sqlconstr)
cnn.CommandTimeout = 900

StrQuery = "SELECT * FROM TABLE_NAME where ModuleName = 'COLUMN_VALUE' AND UserName <> 'COLUMN_VALUE'"

rst.Open StrQuery, cnn

Sheets(3).Range("A2").CopyFromRecordset rst

Dim counter As Long

For counter = 0 To rst.Fields.count - 1
  
  Sheets(3).Range("A1").Offset(, counter).Value = rst.Fields(counter).Name
  
Next

End Sub
