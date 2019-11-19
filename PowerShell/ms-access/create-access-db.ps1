$curDir = get-location

$conn_acc = new-object System.Data.OleDb.OleDbConnection( `
   "Provider=Microsoft.ACE.OLEDB.12.0;" +
   "Data Source=$($curDir)\$($curDir)\new.accdb")

  

$conn_acc
