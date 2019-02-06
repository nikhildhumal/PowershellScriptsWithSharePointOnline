
 
#Provide the Site URL
$SiteUrl = "https://gcpat.sharepoint.com/teams/Quality/GPC-Source-SCC" 

#Provide the First ID and Last ID
$LastItemId=8203
$NextItem=13203
 
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

# Created the Plant Array
$PlantArray=@("Ajax, ON 4202","Augusta, GA  4837", "Calhoun, GA 4622", "Cambridge, MA  4001 R&D","Chicago, IL (65th)", "Chicago, IL (65th)", "Clearfield, UT 4851", "Crystal Warehouse - MA 4052","Denver, CO 4997 T2T","Edmonton, Albt 4205", "El Paso, TX 4973 PBS","Halifax, NS (Storage) 4200 Storage","Houston, TX  4839", "Irondale, AL 4619", "Jackson, MS 4939 PBS","Lithonia, GA 4828", "Los Angeles, CA 4848", "Mansfield, MA 4915 T2T","Milwaukee, WI 4698 Sales Only","Milwaukee, WI 4698 Sales Only","Montreal, QB 4201", "Montreal, QB 4201", "Mt. Pleasant, TN 4046", "North Bergen, NJ 4847", "North Bergen, NJ 4847", "Olathe, KS 4947 PBS","Phoenix, AZ 4821 Toller","Phoenix, AZ 4821 Toller","Pompano Beach, FL 4838 Toller","Pompano Beach, FL 4838 Toller","Pompano Beach, FL 4838 Toller","Puerto Rico  4909", "Santa Ana, CA 4676", "Somerset, NJ 4663 Warehouse","Vancouver, BC 4206 Warehouse","Watermark 4211 Subcontracting","West Chester, OH (Verfi) 4070", "West Chester, OH (Verfi) 4070", "Winnipeg, SK 4203", "Zellwood, FL 4835", "Zellwood, FL 4835", "Archerfield, Australia 4104", "Archerfield, Australia 4104", "Atsugi, Japan 4195", "Balakong, Malaysia 4160", "Balakong, Malaysia 4160", "Balakong, Malaysia 4160", "Balakong, Malaysia 4160", "Bangpoo, Thailand 4170", "Bangpoo, Thailand 4170", "Canlubang, Philippines 4145", "Canlubang, Philippines 4145", "Canning Vale, Australia 4103", "Canning Vale, Australia 4103", "Canning Vale, Australia 4103", "Canning Vale, Australia 4103", "Canning Vale, Australia 4103", "Chongquig, China 4198", "Chongquig, China 4198", "Cikarang, Indonesia 4165", "Cikarang, Indonesia 4165", "Epping, Australia 4101", "Epping, Australia 4101", "Epping, Australia 4101", "Epping, Australia 4101", "Ezhou, China 4189", "HaiDuong, Vietnam 4199", "HaiDuong, Vietnam 4199", "HaiDuong, Vietnam 4199", "HaiDuong, Vietnam 4199", "HocMon, Vietnam 4171", "HocMon, Vietnam 4171", "HocMon, Vietnam 4171", "Holden Hill, Australia 4106", "Hong Kong (Fanling) 4175" ,"Hong Kong (Fanling) 4175" ,"Inchon, Korea 4180" ,"Minhang, China  4190" ,"Minhang, China  4190" ,"Porirua, New Zealand 4110" ,"Regents Park Estate, Australia 4102" ,"Regents Park Estate, Australia 4102" ,"Singapore (Jurong) 4150" ,"Singapore (Jurong) 4150" ,"Taiwan 4185  - Closed","Townsville, Australia 4105" ,"Townsville, Australia 4105" ,"XiQing, China (Tianjin) 4192" ,"XiQing, China (Tianjin) 4192" ,"XiQing, China (Tianjin) 4192" ,"XiQing, China (Tianjin) 4192" ,"ZengCheng, China 4191" ,"ZengCheng, China 4191" ,"Bahia, Brazil 4442" ,"Bahia, Brazil 4442" ,"Bahia, Brazil 4442" ,"Bahia, Brazil 4442" ,"Bogota, Colombia 4420" ,"Cartagena, Colombia 4481 Warehouse","Duque de Caxias, Brazil 4443" ,"Duque de Caxias, Brazil 4443" ,"Duque de Caxias, Brazil 4443" ,"Lampa, Chile 4463" ,"Lampa, Chile 4463" ,"Lima, Peru 4471 GCP Intl. Sucursal - Cons","Lima, Peru 4471 GCP Intl. Sucursal - Cons","Panama 4480" ,"Quilmes, Argentina 4450" ,"Receife, Brazil 4444 SCC" ,"Receife, Brazil 4444 SCC","Receife, Brazil 4444 SCC" ,"Santiago Tianguistenco, Mexico 4431" ,"Sorocaba, Brazil 4440" ,"Valencia, Venzuela 4411" ,"Barcelona, Spain 4334 Warehouse","Bellville, South Africa 4347" ,"Bellville, South Africa 4347" ,"Chennai, India 4196" ,"Dammam, KSA 4391" ,"Delhi, India 4176" ,"Dubai, UAE 4390" ,"Dukinfield, GB 4328" ,"Dukinfield, GB 4328" ,"Epernon, France 4315 Warehouse","Essen, Germany 4362" ,"Heist-op-den-Berg, Belgium (DeNeef) 4372" ,"Heist-op-den-Berg, Belgium (DeNeef) 4372" ,"Helsingborg, Sweden 4340 Warehouse","Helsingborg, Sweden 4340 Warehouse","Jeddah, KSA 4392" ,"Larnaud, France 4317" ,"Luegde, Germany 4361" ,"Passirana, Italy 4325" ,"Slough, UK 4322 Office","Spartan, South Africa 4348" ,"Tuzla, Turkey 4380" ,"Tuzla, Turkey 4380" ,"Widnes, UK 4321" ,"Tolled, Traded Good, Other","Tolled, Traded Good, Other","Archive, plant not active")

Write-Host $PlantArray.Length

# Created the Search Array
$SearchArray =@("Ajax","Augusta","Calhoun","Cambridge","Bedford Park","Chicago","Clearfield","Crystal","Denver","Edmonton","Paso","Halifax","Houston","Irondale","Jackson","Lithonia","Angeles","Mansfield","Enoree","Mil","Montreal","Mont ","Pleasant","North","Bergen","Olathe","Phoen ","Phoenix","Pomp ","Pompano","Pbeach","Puerto","Santa","Somerset","Vancouver","Watermark","West","Chester","Winnipeg","Zellwood","Zell ","Archerfield","Arch ","Atsugi","Balakong","KL","Malaysia","Kuala Lumpur","Bangpoo","Bang ","Canlubang","Can ","Canningvale","Canning Vale","Cann ","Kewdale","Kew ","Chong ","Chongquig","Cikarang","Cik ","Epping","Epp ","Fawkner","Fawk ","Ezhou","Hanoi","HaiDuong","Han ","Hai ","HocMon","HoChiMinh","Chi ","Holden","Hong Kong","Fanling","Inchon","Minhang","Min ","Porirua","Regents","Reg ","Singapore","Jurong","Chung Li","Town ","Townsville","Tianjin","XiQing","TanGu","Tangu","Zeng","Zen ","Bahia","Filho ","S.Filho","Simones","Bogota","Cartagena","Duque","Caxias","Rio","Lampa","Chile","Lima","Peru","Panama","Quilmes","Igarassu","Recife","Receife","Santiago","Sorocaba","Valencia","Barcelona","Bellville","Bell ","Chennai","Dammam","Delhi","Dubai","Dukin ","Dukinfield","Epernon","Essen","Widnes","Tuzla","Turkey","Spartan","Slough","Passirana","Luegde","Larnaud","Jeddah","Helsingborg","Helsing ","Heist","Belgium","Tolled", "Traded","San Boi")

Write-Host $SearchArray.Length

# Loop throgh each and every array and Update the Region value accordingly
foreach ($File in $result) 
{
	if($File.FileSystemObjectType -eq "File")
	{
		For ($i=0; $i -lt $SearchArray.Length; $i++) 
		{     
			if($($File["FileLeafRef"]) -match $SearchArray[$i])
			{							
					$File["Plant"] =$PlantArray[$i] 	 
					$File.update()	
					$clientContext.ExecuteQuery()
					Write-Host "Updating" $($File["FileLeafRef"])
					Write-Host $($File["ID"])				
					"FileName: $($File["FileLeafRef"]), ID:$($File["ID"])" | Add-Content -Path 'E:\Logs Folder\SCCPlantFieldsFromFile.txt'
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
    "Error: $SPOConnectionException" | Add-Content -Path 'E:\Logs Folder\SCCPlantFieldsFromFile.txt'		
    Break
}
