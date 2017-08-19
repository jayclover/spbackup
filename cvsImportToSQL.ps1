Param([string]$SQLServer = $(throw "missing SQLServer as a first argument!"), [string]$SQLDBName = $(throw "missing SQLDBName as a first argument!"), [string]$tableName = $(throw "missing cvsFile as a first argument!"))

#time for starting scrpit
$start = Get-Date

$ErrorActionPreference = "Stop"

$scriptRootPath = $MyInvocation.MyCommand.Path
$configRootDir = Split-Path $scriptRootPath | Split-Path -Parent
write-host "it is $configRootDir"


#Parameterized query
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
 
 return $sqlcmd

}

#call excel and query service to generate csv with all sharepoint data
$xl = New-Object -C Excel.Application -vb:$false
$xl.DisplayAlerts = $False
$queryFilePath = $configRootDir + '\spbackup\query.iqy'
$iqy = $xl.Workbooks.Open($queryFilePath)
$cvsFilePath = $configRootDir + '\spbackup\listDataCSV.csv'
$iqy.SaveAs($cvsFilePath, 6)

#Close the excel
$xl.quit()
$confirm = $false
Get-Process Excel | kill -Confirm:$confirm


#Create the SQL Connection Object
Write-Verbose "Creating SQL Connection"
$SQLConn = New-Object System.Data.SqlClient.SqlConnection("Data Source=$SQLServer; Initial Catalog=$SQLDBName; Integrated Security=SSPI")
$SQLConn.Open()

# Quit if the SQL connection didn't open properly.
if ($SQLConn.State -ne [Data.ConnectionState]::Open) {
    write-host -Fore Red "Connection to DB is not open."
    Exit
}


#read the csv file, the encoding Default is a must here to avoid encoding issue
$spDatas = import-csv $cvsFilePath -Encoding Default

#detet the row from SQL which existed in csv for updating
foreach($spData in $spDatas) {
	$removeTitle = $spData.'Title'
	
	If ($removeTitle) {
		$deleteQuery = "delete from $tableName where [Title]=@Title"

		Write-Host -Fore Yellow "Remove the legacy row which title is: $removeTitle"
		exec-query $deleteQuery -parameter @{'Title'= $spData.'Title'} -sqlconn $sqlconn	
	} else {
		Write-Host -Fore blue "This row didn't have value for title"
	}

}

# Pick up all column name from the csv with sharepoint data
$spMember = $spDatas |Get-member
$spMemberName = $spMember.name

# write all data from csv with sharepoint data into SQL
$spDatas | % {
	
	$updateTitle = $_.'Title'
	write-host -Fore Green "Update the row for title: $updateTitle"
	
	$columnListPath = $configRootDir + '\spbackup\columnlist.csv'
	$columns = Import-Csv $columnListPath
	$parameter=@{}
	$queryColumnname = "("
	$queryColumncode = "("
	foreach($column in $columns) {
		$columnName = $column.'column name'
		$columncode =$column.'column code'

		#checking if column name is matched in the csv with sharepoint data and the csv with column dictionary
		if ($spMemberName.Contains($columnName)){
			$parameter+=@{$columncode = $($_.$columnName)}
		} else {
		    write-host -Fore Red "Column table didn't match all the field"
			Exit
		}					
		$queryColumnname += "[" + $columnName +"]"+","
		$queryColumncode += "@" + $columncode +","
	}
	$queryColumnname = $queryColumnname.Trim(",")
	$queryColumnname = $queryColumnname + ")"
	$queryColumncode = $queryColumncode.Trim(",")
	$queryColumncode = $queryColumncode + ")"
	
	$insertQuery = "INSERT INTO $tableName $queryColumnname VALUES $queryColumncode"
	exec-query $insertQuery -parameter $parameter -sqlconn $sqlconn	
	
 }
  
#Close
$SQLConn.Close()

$end = Get-Date
Write-Host -ForegroundColor Green ('Total Runtime: ' + ($end - $start).TotalSeconds)

Write-Host "Done."
exit
