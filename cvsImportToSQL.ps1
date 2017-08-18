Param([string]$SQLServer = $(throw "missing SQLServer as a first argument!"), [string]$SQLDBName = $(throw "missing SQLDBName as a first argument!"), [string]$tableName = $(throw "missing cvsFile as a first argument!"), [string]$cvsFilePath = $(throw "missing cvsFile as a first argument!"))


function exec-query( $query,$parameters=@{},$SQLConn,$timeout=30,[switch]$help){
 if ($help){
 $msg = @"
Execute a sql statement.  Parameters are allowed.
Input parameters should be a dictionary of parameter names and values.
Return value will usually be a list of datarows.
"@
 Write-Host $msg
 return
 }
 $SQLcmd=new-object system.Data.SqlClient.SqlCommand($query,$SQLConn)
 $SQLcmd.CommandTimeout=$timeout
 foreach($p in $parameters.Keys){
 [Void] $SQLcmd.Parameters.AddWithValue("@$p",$parameters[$p])
 }
 $ds=New-Object system.Data.DataSet
 $da=New-Object system.Data.SqlClient.SqlDataAdapter($SQLcmd)
 $da.fill($ds) | Out-Null
 
 return $ds

}


#Create the SQL Connection Object
Write-Verbose "Creating SQL Connection"
$SQLConn = New-Object System.Data.SqlClient.SqlConnection("Data Source=$SQLServer; Initial Catalog=$SQLDBName; Integrated Security=SSPI")
$SQLConn.Open()

# Quit if the SQL connection didn't open properly.
if ($SQLConn.State -ne [Data.ConnectionState]::Open) {
    write-host -Fore Red "Connection to DB is not open."
    Exit
}

$spDatas = Import-Csv $cvsFilePath

foreach($spData in $spDatas) {
	$Title1 = $spData.'Title'
	
	If ($Title1) {
		$deleteQuery = "delete from $tableName where [Title]=@Title"

		Write-Host -Fore Yellow "Remove the legacy row which title is: $Title1"
		exec-query $deleteQuery -parameter @{'Title'= $spData.'Title'} -sqlconn $sqlconn	
	} else {
		Write-Host -Fore white "This row didn't have value for title"
	}

}

Import-Csv $cvsFilePath | % {
	
	$title1 = $_.'Title'
	write-host -Fore Green "Update the row for title: $title1"

	$parameter=@{}
	$parameter+=@{'request' = $($_.'Request Stage')}
	$parameter+=@{'title' = $($_.'Title')}
	
	$insertQuery = "INSERT INTO $tableName ([Request Stage],[Title]) VALUES (@request,@title)"
	exec-query $insertQuery -parameter $parameter -sqlconn $sqlconn	
	
 }
  
#Close
$SQLConn.Close()

