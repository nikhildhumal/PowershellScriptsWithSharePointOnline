
# Provide the Site URL
$SiteUrl = "https://gcpat.sharepoint.com/teams/Quality/GPC-Source-SCC"       

#Provide the First ID and Last ID
$LastItemId=0
$NextItem=5000


# Provide the Username and Password
$UserName = "nikhil.dhumal@gcpat.com" # Read-Host -Prompt "Enter User Name"
$SecurePassword = Read-Host -Prompt "Enter password" -AsSecureString

Try
{
    # Import the Sharepoint client object model dlls
    Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll"
    Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"
    
    
    $ctx = New-Object Microsoft.SharePoint.Client.ClientContext($SiteUrl)
    $credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($UserName, $SecurePassword)
    $ctx.Credentials = $credentials
    $lists = $ctx.web.Lists
# Provide the Document Library Name
$list = $lists.GetByTitle("Formulas")
# Caml query that traverse through all items under document Library
<#
$query =[Microsoft.SharePoint.Client.CamlQuery]::CreateAllItemsQuery(10000)
$result = $list.GetItems($query)
$ctx.Load($lists)
$ctx.Load($result)
$ctx.ExecuteQuery()
#>

$qCommand = @"
<View Scope="RecursiveAll">
    <Query>
    <Where>
    <And>
      <Geq>
         <FieldRef Name='ID' />
         <Value Type='Counter'>$LastItemId</Value>
      </Geq>
      <Leq>
      <FieldRef Name='ID' />
         <Value Type='Counter'>$NextItem</Value>
      
      </Leq>
      </And>
   </Where>
        <OrderBy><FieldRef Name='ID' Ascending='TRUE'/></OrderBy>
    </Query>
    <RowLimit Paged="TRUE">5000</RowLimit>
</View>
"@

## Page Position
$position = $null
 
## All Items
$allItems = @()

    $camlQuery = New-Object Microsoft.SharePoint.Client.CamlQuery
    $camlQuery.ListItemCollectionPosition = $position
    $camlQuery.ViewXml = $qCommand
 ## Executing the query
    $result = $list.GetItems($camlQuery)
    $ctx.Load($result)
    $ctx.ExecuteQuery()




# Loop throgh each and every files under document library if filename contains any special character like "-" then split it and update field values
foreach ($File in $result) 
{
		if($File.FileSystemObjectType -eq "File")
		{ 
			$File["DocType"] = "Formula"
            $File["Business"] = "SCC"
			$File.update()	
			$ctx.ExecuteQuery()
			Write-Host "Updating" $($File["FileLeafRef"])
			Write-Host $($File["ID"])
			"FileName: $($File["FileLeafRef"]), ID:$($File["ID"])" | Add-Content -Path 'E:\Logs Folder\SCCTraded & Tolled GoodsLogs.txt'	
		}
	
	<# 
	if($File.FileSystemObjectType -eq "File")
		{
			if($($File["FileLeafRef"]) -match "-")
			{
				$InitialFilename = $($File["FileLeafRef"]).split("-")[0]
				$MiddleFilename = $($File["FileLeafRef"]).split("-")[1]
				$LastFilename = $($File["FileLeafRef"]).split("-")[2].split(".")[0]	
				$File["FileName"] = $InitialFilename   
				$File["Source"] = $MiddleFilename	 
				$File.update()	
				$ctx.ExecuteQuery()
				Write-Host "Updating" $($File["FileLeafRef"])
				Write-Host $($File["ID"])
			} 
		}
	#>

}

	  
}
Catch
{
    $SPOConnectionException = $_.Exception.Message
    Write-Host ""
    Write-Host "Error:" $SPOConnectionException -ForegroundColor Red
    Write-Host ""
    "Error: $SPOConnectionException" | Add-Content -Path 'E:\Logs Folder\SCCTraded & Tolled GoodsLogs.txt'	
    Break
}

