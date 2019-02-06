
# Provide the Site URL
$SiteUrl = "https://gcpat.sharepoint.com/teams/Quality/GPC-Source-SCC"       

#Provide the First ID and Last ID
$LastItemId=5000
$NextItem=10000

# Provide the UserName and Password
$UserName = "nikhil.dhumal@gcpat.com" # Read-Host -Prompt "Enter User Name"
$SecurePassword = Read-Host -Prompt "Enter password" -AsSecureString

Try
{
    # Import the Sharepoint Client object model dll
    Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll"
    Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"
    
    
    $ctx = New-Object Microsoft.SharePoint.Client.ClientContext($SiteUrl)
    $credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($UserName, $SecurePassword)
    $ctx.Credentials = $credentials
    $lists = $ctx.Web.Lists
$ctx.load($lists)
$ctx.ExecuteQuery()
$list = $lists.GetByTitle("Formulas")
$ctx.load($list)
$ctx.ExecuteQuery()
Write-Host $web.ServerRelativeUrl
# Caml Query that traverse through each and every folder under document library
<#
$q = New-Object Microsoft.SharePoint.Client.CamlQuery
$q.ViewXml = '<View Scope="RecursiveAll"><Query><Where><Eq><FieldRef Name="FSObjType" /><Value Type="Integer">0</Value></Eq></Where><OrderBy><FieldRef Name="ID" /></OrderBy></Query></View>'
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
    $camlQuery.FolderServerRelativeUrl= "/teams/Quality/GPC-Source-SCC/Formulas/SCC/Decorative"
 ## Executing the query
    $items = $list.GetItems($camlQuery)
    $ctx.Load($items)
    $ctx.ExecuteQuery()

        
				foreach($Filename in $items)
      				{ 
					    if($Filename.FileSystemObjectType -eq "File")
					    {
        	 				$Filename["ProductLine"] ="Decorative"
        	 				$Filename.update();
    		 				$ctx.ExecuteQuery()
                            Write-Host "Updating" $($Filename["FileLeafRef"])
                            Write-Host $($Filename["ID"])
			                "FileName: $($Filename["FileLeafRef"]), ID:$($Filename["ID"])" | Add-Content -Path 'E:\Logs Folder\SCCConcrete AdmixturesFieldsLogs.txt'		
					    }
				    }
			
		       
	  
}
Catch
{
    $SPOConnectionException = $_.Exception.Message
    Write-Host ""
    Write-Host "Error:" $SPOConnectionException -ForegroundColor Red
    Write-Host ""
    "Error: $SPOConnectionException" | Add-Content -Path 'E:\Logs Folder\SCCConcrete AdmixturesFieldsLogs.txt'			
    Break
}

