Param([string]$SQLServer = $(throw "missing SQLServer as a first argument!"), [string]$SQLDBName = $(throw "missing SQLDBName as a first argument!"), [string]$tableName = $(throw "missing cvsFile as a first argument!"), [string]$cvsFilePath = $(throw "missing cvsFile as a first argument!"))


#Create the SQL Connection Object
Write-Verbose "Creating SQL Connection"
$SQLConn = New-Object System.Data.SqlClient.SqlConnection("Data Source=$SQLServer; Initial Catalog=$SQLDBName; Integrated Security=SSPI")
$SQLConn.Open()
#Create the SQL Command Object, to work with the Database
$SQLCmd = $SQLConn.CreateCommand()




$spDatas = Import-Csv $cvsFilePath

foreach($spData in $spDatas) {
	$spRequestID = $spData.'requests list Linked Id'
	if ($spRequestID -ne "1321") {
		$deleteQuery = "delete from $tableName where [Requests List Linked ID]=$spRequestID"

		$SQLCmd.CommandText = $deleteQuery
		Write-Host -Fore Yellow "Remove the legacy row in table"
  
		#Execute Query
		$SQLCmd.ExecuteNonQuery() | Out-Null 
	}
}

$rowNo = $NULL
Import-Csv $cvsFilePath | % {

	#$insertQuery = "INSERT INTO $tableName  VALUES ('$($_.'Request Stage')','$($_.'Title')','$($_.'Created By')','$($_.'Created')','$($_.'Request Type')','$($_.'Purchase Order Total')','$($_.'Expense Owner or EXT Sponsor')','$($_.'Org')','$($_.'PO Adjust'),'$($_.'Add Resource Title')','$($_.'Resource 1 PA')','$($_.'Resource 2 PA')','$($_.'Resource 3 PA')','$($_.'Resource 4 PA')','$($_.'Resource 5 PA')','$($_.'Resource 1 RC')','$($_.'Resource 2 RC')','$($_.'Resource 3 RC')','$($_.'Resource 4 RC')','$($_.'Resource 5 RC')','$($_.'group1')','$($_.'PO Amt Original')','$($_.'PO Amt Extended')','$($_.'field1')','$($_.'Extended PO Number')','$($_.'Requests List Linked ID')','$($_.'Supplier Number')','$($_.'Milestone 1 IO')','$($_.'Milestone 2 IO')','$($_.'Milestone 3 IO')','$($_.'Milestone 4 IO')','$($_.'Milestone 5 IO')','$($_.'Milestone 6 IO')','$($_.'ExtPO Org')','$($_.'EXT Bulk Submit')','$($_.'Expedite')','$($_.'ExpediteR')','$($_.'Threshold Approved')','$($_.'MyDash - Request Process')','$($_.'MyDash - Get Org')','$($_.'MyDash - Send to CRM')','$($_.'PO NUMBER')','$($_.'Item Type')','$($_.'Path')')"
	$rowNo++
	Write-Host -Fore Green "Updateing row# $rowNo"
	$values = @()
	$SQLCmd.CommandText = "INSERT INTO $tableName  VALUES ('$($_.'Request Stage')','$($_.'Title')','$($_.'Created By')','$($_.'Created')','$($_.'Request Type')','$($_.'Purchase Order Total')','$($_.'Expense Owner or EXT Sponsor')','$($_.'Org')','$($_.'PO Adjust'),'$($_.'Add Resource Title')','$($_.'Resource 1 PA')','$($_.'Resource 2 PA')','$($_.'Resource 3 PA')','$($_.'Resource 4 PA')','$($_.'Resource 5 PA')','$($_.'Resource 1 RC')','$($_.'Resource 2 RC')','$($_.'Resource 3 RC')','$($_.'Resource 4 RC')','$($_.'Resource 5 RC')','$($_.'group1')','$($_.'PO Amt Original')','$($_.'PO Amt Extended')','$($_.'field1')','$($_.'Extended PO Number')','$($_.'Requests List Linked ID')','$($_.'Supplier Number')','$($_.'Milestone 1 IO')','$($_.'Milestone 2 IO')','$($_.'Milestone 3 IO')','$($_.'Milestone 4 IO')','$($_.'Milestone 5 IO')','$($_.'Milestone 6 IO')','$($_.'ExtPO Org')','$($_.'EXT Bulk Submit')','$($_.'Expedite')','$($_.'ExpediteR')','$($_.'Threshold Approved')','$($_.'MyDash - Request Process')','$($_.'MyDash - Get Org')','$($_.'MyDash - Send to CRM')','$($_.'PO NUMBER')','$($_.'Item Type')','$($_.'Path')')"
	for($i = 0; $i -lt @($values).Count; $i++) {
		$SQLCmd.Parameters[$i].Value = $values[$i] | Out-Null
	}
  #Execute Query
  $SQLCmd.ExecuteNonQuery() | Out-Null 
 }
  
  
#Close
$SQLConn.Close()
