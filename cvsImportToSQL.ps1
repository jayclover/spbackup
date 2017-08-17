Param([string]$SQLServer = $(throw "missing SQLServer as a first argument!"), [string]$SQLDBName = $(throw "missing SQLDBName as a first argument!"), [string]$tableName = $(throw "missing cvsFile as a first argument!"), [string]$cvsFilePath = $(throw "missing cvsFile as a first argument!"))


#Create the SQL Connection Object
Write-Verbose "Creating SQL Connection"
$SQLConn = New-Object System.Data.SqlClient.SqlConnection("Data Source=$SQLServer; Initial Catalog=$SQLDBName; Integrated Security=SSPI")
$SQLConn.Open()
#Create the SQL Command Object, to work with the Database
$SQLCmd = $SQLConn.CreateCommand()


$insertQuery = "INSERT INTO $tableName  VALUES ('$($_.'Request Stage')','$($_.Title)','$($_.’Requests List Linked ID‘)')"
$selectQuery = "select * from $tableName"
$deleteQuery = "delete from $tableName where 'Requests List Linked ID' = $spRequestID"
#Handle the query with SQLCommand Object
$SQLCmd.CommandText = $selectQuery

$SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
$SqlAdapter.SelectCommand = $SQLCmd
$DataSet = New-Object System.Data.DataSet
$nRecs = $SqlAdapter.Fill($DataSet)
$nRecs | Out-Null

#Populate Hash Table
$objTable = $DataSet.Tables[0]

#Export Hash Table to CSV File
$legacyTableVaule = "legacyTableVaule.csv"
$objTable | Export-CSV $legacyTableVaule


$spDatas = Import-Csv $cvsFilePath
$sqlDatas = Import-Csv $legacyTableVaule

foreach($spData in $spDatas) {
	$spRequestID = $spData.'requests list Linked Id'
	foreach($sqlData in $sqlDatas) {
		$sqlRequestID = $sqlData.'requests list Linked Id'
			if ($spRequestID -eq $sqlRequestID) {
				$SQLCmd.CommandText = $deleteQuery
				Write-Host -Fore Green "Remove the legacy row in table"
  
				#Execute Query
				$SQLCmd.ExecuteNonQuery() | Out-Null 
			}
		$SQLCmd.CommandText = $insertQuery
		
		Write-Host -Fore Green "Updateing Table"
		#Execute Query
		$SQLCmd.ExecuteNonQuery() | Out-Null 
	
	}
}

#Import-Csv $cvsFilePath | % {

  #Write-Host -Fore Green "Updateing Table"
  
  #Execute Query
  #$SQLCmd.ExecuteNonQuery() | Out-Null 
  #}
  
  
#Close
$SQLConn.Close()
