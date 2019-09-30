$curDir = get-location

$conn_xlsx = new-object System.Data.OleDb.OleDbConnection( `
   "Provider=Microsoft.ACE.OLEDB.12.0;" +
   "Data Source=$($curDir)\data.xlsx;"  +
   "Extended Properties=""Excel 12.0 Xml""")

$query = new-object System.Data.OleDb.OleDbCommand("
  select
     en.id,
     en.val as val_en,
     de.val as val_de
  from
     [en$]  en left join
     [de$]  de on en.id = de.id
")

$query.connection = $conn_xlsx

$conn_xlsx.open()

$reader = $query.ExecuteReader()

while ($reader.read()) {
  echo("{0, 2}: {1, -5}  {2, -5}" -f $reader['id'], $reader['val_en'], $reader['val_de'])
}
