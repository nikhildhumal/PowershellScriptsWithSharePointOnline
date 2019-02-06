# this method find the Folder Grace and update all Filename field in document Library as Grace value
function GetChildFolders($folder) 
{
	$ctx.Load($folder)
	$ctx.ExecuteQuery()
        $ctx.Load($folder.Folders)
        $ctx.ExecuteQuery()

        foreach($Subfolder in $folder.Folders)
        {
          	if($Subfolder.Name -ne "Forms")
          	{		
				foreach($Filename in $items)
      				{ 
					    if($Filename.FileSystemObjectType -eq "File")
					    {
        	 				$Filename["Business"] ="SCC"
                            $Filename["ProductLine"] ="Pieri"
                            $Filename["Plant"] ="Larnaud, France 4317"
        	 				$Filename.update();
    		 				$ctx.ExecuteQuery()
                            Write-Host "Updating" $($Filename["FileLeafRef"])
                            Write-Host $($Filename["ID"])
			    "FileName: $($Filename["FileLeafRef"]), ID:$($Filename["ID"])" | Add-Content -Path 'C:\Users\amit_dahotre\Desktop\UpdateAllFieldsRecursivelyLogs.txt'		
					    }
				    }
	
        	}
	GetChildFolders $Subfolder
	}
}


# Provide the Site URL
$SiteUrl = "https://gcpat.sharepoint.com/teams/QA-Quality/GPC-Source"       

#Provide the First ID and Last ID
$LastItemId=0
$NextItem=5000

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
    $lists = $ctx.web.Lists
$ctx.load($lists)
$ctx.ExecuteQuery()
$list = $lists.GetByTitle("Formulas")
$ctx.load($list)
$ctx.ExecuteQuery()
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
 ## Executing the query
    $items = $list.GetItems($camlQuery)
    $ctx.Load($items)
    $ctx.ExecuteQuery()

$ctx.Load($list.RootFolder.Folders)
           $ctx.ExecuteQuery()
           write-host 'RootFolder.Folders.Count:', $list.RootFolder.Folders.Count
           
  foreach ($foldername in $list.RootFolder.Folders)
           {
		if($foldername.Name -ne “Forms”)
		{
              		# If folder name is grace then update the filename as Grace
			if($foldername.name -eq "Pieri Products")
			{
				foreach($Filename in $items)
      				{ 
					if($Filename.FileSystemObjectType -eq "File")
					{
        	 				$Filename["Business"] ="SCC"
                            $Filename["ProductLine"] ="Pieri"
                            $Filename["Plant"] ="Larnaud, France 4317"    
        	 				$Filename.update();
    		 				$ctx.ExecuteQuery()
                            Write-Host "Updating" $($Filename["FileLeafRef"])
                            Write-Host $($Filename["ID"])
			    "FileName: $($Filename["FileLeafRef"]), ID:$($Filename["ID"])" | Add-Content -Path 'C:\Users\amit_dahotre\Desktop\UpdateAllFieldsRecursivelyLogs.txt'		
					}
				}
			}
			# call the Getchildfolders method
			GetChildFolders $foldername 
		}
           }       
	  
}
Catch
{
    $SPOConnectionException = $_.Exception.Message
    Write-Host ""
    Write-Host "Error:" $SPOConnectionException -ForegroundColor Red
    Write-Host ""
    "Error: $SPOConnectionException" | Add-Content -Path 'C:\Users\amit_dahotre\Desktop\UpdateAllFieldsRecursivelyLogs.txt'			
    Break
}

