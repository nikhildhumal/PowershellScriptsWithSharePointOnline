
# Provide the Site Url
$SiteUrl = "https://gcpat.sharepoint.com/teams/QA-Quality/GPC-Source"       

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
# Provide the Folder Relative Path    
$folderRelativeUrl ="/Formulas/Pieri Products/Pieri Products" 
$web = $ctx.Web 
$ctx.Load($web)
$ctx.ExecuteQuery()

write-host $web.ServerRelativeUrl

$folder = $web.GetFolderByServerRelativeUrl($web.ServerRelativeUrl + $folderRelativeUrl)


$ctx.Load($folder)
$ctx.ExecuteQuery()
Write-Host "Files from " $folder.Name
$files = $folder.Files
$ctx.Load($folder.Files)
$ctx.Load($folder.Folders)
$ctx.ExecuteQuery()
# Loop through each files under Clone Folder and Update the Source Field as Clone which is name of the Folder
foreach($file in $files)
{ 	
        Write-Host $file.Name
        $file.Properties["Business"] = "SCC"
        $file.Properties["Plant"] = "Larnaud, France 4317"
        $file.Properties["ProductLine"] = "Pieri"
        $file.update();
	    $ctx.ExecuteQuery()
	"FileName: $($File["FileLeafRef"]), ID:$($File["ID"])" | Add-Content -Path 'C:\Users\amit_dahotre\Desktop\UpdateFieldsByreadingFolderRalativePathLogs.txt'	
   
}
 
}
Catch
{
    $SPOConnectionException = $_.Exception.Message
    Write-Host ""
    Write-Host "Error:" $SPOConnectionException -ForegroundColor Red
    Write-Host ""
   "Error: $($File["FileLeafRef"])" | Add-Content -Path 'C:\Users\amit_dahotre\Desktop\UpdateFieldsByreadingFolderRalativePathLogs.txt'	
    Break
}

