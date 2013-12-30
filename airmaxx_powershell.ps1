#######################

## runs okay when cut and pasted into powershell launched from mssql server management studio
## 

## Exports Blobs found in table 'DocAttach' to file system as zip files         
## with GetBytes-Stream.         
# Configuration data            
$Server = "(local)";         # SQL Server Instance.            
$Database = "some-database";            
# $Dest = "C:\path\to\somewhere";         # Path to export to.            
$Dest = "G:\airmaxx\export\";         # Dump to external drive            

$bufferSize = 8192;               # Stream buffer size in bytes.            
# Select-Statement for name & blob            
# with filter.            
$Sql = "SELECT [ID]
              ,[Document]
	,[Extension]
        FROM dbo.DocAttach";            
            
# Open ADO.NET Connection            
$con = New-Object Data.SqlClient.SqlConnection;            
$con.ConnectionString = "Data Source=$Server;" +             
                        "Integrated Security=True;" +            
                        "Initial Catalog=$Database";            
$con.Open();            
            
# New Command and Reader            
$cmd = New-Object Data.SqlClient.SqlCommand $Sql, $con;            
$rd = $cmd.ExecuteReader();            
            
# Create a byte array for the stream.            
$out = [array]::CreateInstance('Byte', $bufferSize)            
            
# Looping through records            
While ($rd.Read())            
{            
    Write-Output ("Exporting: {0}" -f $rd.GetString(0) ); 
    #$nugget = $rd.GetString(1);
    # New BinaryWriter            
    $fs = New-Object System.IO.FileStream ($Dest + $rd.GetString(0) + '.zip'), Create, Write;            
    $bw = New-Object System.IO.BinaryWriter $fs;            
               
    $start = 0;            
    # Read first byte stream            
    $received = $rd.GetBytes(1, $start, $out, 0, $bufferSize - 1);            
    While ($received -gt 0)            
    {            
       $bw.Write($out, 0, $received);            
       $bw.Flush();            
       $start += $received;            
       # Read next byte stream            
       $received = $rd.GetBytes(1, $start, $out, 0, $bufferSize - 1);            
    }            
            
    $bw.Close();            
    $fs.Close();            
}            
            
# Closing & Disposing all objects            
$fs.Dispose();            
$rd.Close();            
$cmd.Dispose();            
$con.Close();            
            
Write-Output ("Finished");

