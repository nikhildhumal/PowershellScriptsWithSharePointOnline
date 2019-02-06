
 
#Provide the Site URL
$SiteUrl = "https://gcpat.sharepoint.com/teams/Quality/GPC-Source-SCC" 

#Provide the First ID and Last ID
$LastItemId=9999
$NextItem=14999
 
# Provide the Username and Password
$UserName = "nikhil.dhumal@gcpat.com" # Read-Host -Prompt "Enter User Name"
$SecurePassword = Read-Host -Prompt "Enter password" -AsSecureString
Try
{
   # Import the Sharepoint client object model dll
    Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll"
         Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"
   
    
   
    $clientContext = New-Object Microsoft.SharePoint.Client.ClientContext($SiteUrl)
    $credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($UserName, $SecurePassword)
    $clientContext.Credentials = $credentials
    $lists = $clientContext.web.Lists
$list = $lists.GetByTitle("Formulas")
# Caml Query traverse throgh all the Folder and Subfolder inside Document Library or particular List
<#
$query =[Microsoft.SharePoint.Client.CamlQuery]::CreateAllItemsQuery(10000)
$result = $list.GetItems($query)
$clientContext.Load($lists)
$clientContext.Load($result)
$clientContext.ExecuteQuery()
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
    $clientContext.Load($result)
    $clientContext.ExecuteQuery()


$NA = "North America"
$AP = "Asia Pacific"
$LA = "Latin America"
$EMEA = "EMEA"
$TOLLEDORTRADED ="Tolled or Traded" 
$ARCHIVE="ARCHIVE"

# Created the Region Array
$NAArray =@("Ajax","Augusta","Calhoun","Cambridge","Bedford Park","Chicago","Clearfield","Crystal","Denver","Edmonton","Paso","Halifax","Houston","Irondale","Jackson","Lithonia","Angeles","Mansfield","Enoree","Mil","Montreal","Mont ","Pleasant","North","Bergen","Olathe","Phoen ","Phoenix","Pomp ","Pompano","Pbeach","Puerto","Santa","Somerset","Vancouver","Watermark","West","Chester","Winnipeg","Zellwood","Zell ")

$APArray = @("Archerfield","Arch ","Atsugi","Balakong","KL","Malaysia","Kuala Lumpur","Bangpoo","Bang ","Canlubang","Can ","Canningvale","Canning Vale","Cann ","Kewdale","Kew ","Chong ","Chongquig","Cikarang","Cik ","Epping","Epp ","Fawkner","Fawk ","Ezhou","Hanoi","HaiDuong","Han ","Hai ","HocMon","HoChiMinh","Chi ","Holden","Hong Kong","Fanling","Inchon","Minhang","Min ","Porirua","Regents","Reg ","Singapore","Jurong","Chung Li","Town ","Townsville","Tianjin","XiQing","TanGu","Tangu","Zeng","Zen ")

$LAArray = @("Bahia","Filho ","S.Filho","Simones","Bogota","Cartagena","Duque","Caxias","Rio","Lampa","Chile","Lima","Peru","Panama","Quilmes","Igarassu","Recife","Receife","Santiago","Sorocaba","Valencia")
 
$EMEAArray =@("Barcelona","Bellville","Bell ","Chennai","Dammam","Delhi","Dubai","Dukin ","Dukinfield","Epernon","Essen","Widnes","Tuzla","Turkey","Spartan","Slough","Passirana","Luegde","Larnaud","Jeddah","Helsingborg","Helsing ","Heist","Belgium")

$TOLLEDORTRADEDArray=@("Tolled", "Traded")

$ARCHIVEArray =@("San Boi")

$NextFile=""

# Loop throgh each and every array and Update the Region value accordingly
foreach ($File in $result) 
{
	if($File.FileSystemObjectType -eq "File")
	{
        $NextFile=""
		For ($i=0; $i -lt $NAArray.Length; $i++) 
		{     
			if($($File["FileLeafRef"]) -match $NAArray[$i])
			{
                
				$File["Region"] = $NA	 
				$File.update()	
				$clientContext.ExecuteQuery()
				Write-Host "Updating" $($File["FileLeafRef"])
				Write-Host $($File["ID"])
				"FileName: $($File["FileLeafRef"]), ID:$($File["ID"])" | Add-Content -Path 'E:\Logs Folder\SCCRegionFieldsFromFile.txt'
                $NextFile = "True"
                break
			}
		}
        if($NextFile -eq "True")
        {
            Continue
        }

	For ($j=0; $j -lt $APArray.Length; $j++) 
	{     
		if($($File["FileLeafRef"]) -match $APArray[$j])
		{
			$File["Region"] = $AP	 
			$File.update()	
			$clientContext.ExecuteQuery()
			Write-Host "Updating" $($File["FileLeafRef"])
			Write-Host $($File["ID"])
			"FileName: $($File["FileLeafRef"]), ID:$($File["ID"])" | Add-Content -Path 'E:\Logs Folder\SCCRegionFieldsFromFile.txt'
            $NextFile = "True"
                break
		}

	}
    if($NextFile -eq "True")
        {
            Continue
        }

	For ($k=0; $k -lt $LAArray.Length; $k++) 
	{     
		if($($File["FileLeafRef"]) -match $LAArray[$k])
		{
			$File["Region"] = $LA	 
			$File.update()	
			$clientContext.ExecuteQuery()
			Write-Host "Updating" $($File["FileLeafRef"])
			Write-Host $($File["ID"])
			"FileName: $($File["FileLeafRef"]), ID:$($File["ID"])" | Add-Content -Path 'E:\Logs Folder\SCCRegionFieldsFromFile.txt'
            $NextFile = "True"
                break
		}
	}
    if($NextFile -eq "True")
        {
            Continue
        }

	For ($l=0; $l -lt $EMEAArray.Length; $l++) 
	{     
		if($($File["FileLeafRef"]) -match $EMEAArray[$l])
		{
			$File["Region"] = $EMEA	 
			$File.update()	
			$clientContext.ExecuteQuery()
			Write-Host "Updating" $($File["FileLeafRef"])
			Write-Host $($File["ID"])
			"FileName: $($File["FileLeafRef"]), ID:$($File["ID"])" | Add-Content -Path 'E:\Logs Folder\SCCRegionFieldsFromFile.txt'
            $NextFile = "True"
                break
		}
	}
    if($NextFile -eq "True")
        {
            Continue
        }
	For ($m=0; $m -lt $TOLLEDORTRADEDArray.Length; $m++) 
	{     
		if($($File["FileLeafRef"]) -match $TOLLEDORTRADEDArray[$m])
		{
			$File["Region"] = $TOLLEDORTRADED	 
			$File.update()	
			$clientContext.ExecuteQuery()
			Write-Host "Updating" $($File["FileLeafRef"])
			Write-Host $($File["ID"])
			"FileName: $($File["FileLeafRef"]), ID:$($File["ID"])" | Add-Content -Path 'E:\Logs Folder\SCCRegionFieldsFromFile.txt'
            $NextFile = "True"
                break
		}
	}
    if($NextFile -eq "True")
        {
            Continue
        }
	For ($n=0; $n -lt $ARCHIVEArray.Length; $n++) 
	{     
		if($($File["FileLeafRef"]) -match $ARCHIVEArray[$n])
		{
			$File["Region"] = $ARCHIVE	 
			$File.update()	
			$clientContext.ExecuteQuery()
			Write-Host "Updating" $($File["FileLeafRef"])
			Write-Host $($File["ID"])
			"FileName: $($File["FileLeafRef"]), ID:$($File["ID"])" | Add-Content -Path 'E:\Logs Folder\SCCRegionFieldsFromFile.txt'
            break
		}
	}

}

}

	    

}
Catch
{
    $SPOConnectionException = $_.Exception.Message
    Write-Host ""
    Write-Host "Error:" $SPOConnectionException -ForegroundColor Red
    Write-Host ""
    "Error: $SPOConnectionException" | Add-Content -Path 'E:\Logs Folder\SCCRegionFieldsFromFile.txt'
    Break
}
