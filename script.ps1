$fileName = "Database1.mdb"
$conn = New-Object System.Data.OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=$filename;Persist Security Info=False")
$cmd=$conn.CreateCommand()
$cmd.CommandText="Select * from customers"
$conn.open()
$rdr = $cmd.ExecuteReader()
$dt = New-Object System.Data.Datatable
$dt.Load($rdr)
$dt


