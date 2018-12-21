$conn = New-Object System.Data.OleDb.OleDbConnection;

#請使用 ORAOLEDB.Oracle 做為新的 Provider
$conn.ConnectionString = "Provider=ORAOLEDB.Oracle;Password=oracle;User ID=system;Data Source={server}";

#使用 MSDAORA.1 會出現 ORA-01019 錯誤
#$conn.ConnectionString = "Provider=MSDAORA.1;Password=oracle;User ID=system;Data Source={server}";

$cmd = New-Object System.Data.OleDb.OleDbCommand
$cmd.Connection = $conn;
$cmd.CommandText = "SELECT SYSDATE FROM DUAL";

$conn.Open();
$value = $cmd.ExecuteScalar();
$value;
$conn.Close();