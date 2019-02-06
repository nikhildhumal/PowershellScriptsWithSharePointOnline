
 
#Provide the Site URL
$SiteUrl = "https://gcpat.sharepoint.com/teams/QA-Quality/GPC-Source" 

#Provide the First ID and Last ID
$LastItemId=0
$NextItem=5000
 
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
$list = $lists.GetByTitle("TestClone")
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
$PlantArray=@("Ajax, ON 4202","Archerfield, Australia 4104","Archerfield, Australia 4104","Atsugi, Japan 4195","Augusta, GA 4837","Bahia, Brazil 4442","Bahia, Brazil 4442","Bahia, Brazil 4442","Bahia, Brazil 4442","Balakong, Malaysia 4160","Balakong, Malaysia 4160","Balakong, Malaysia 4160","Balakong, Malaysia 4160","Bangpoo, Thailand 4170","Bangpoo, Thailand 4170","Barcelona, Spain 4334 Warehouse","Bellville, South Africa 4347","Bellville, South Africa 4347","Bogota, Colombia 4420","Calhoun, GA 4622","Cambridge, MA 4001 R&D","Canlubang, Philippines 4145","Canlubang, Philippines 4145","Canning Vale, Australia 4103","Canning Vale, Australia 4103","Canning Vale, Australia 4103","Canning Vale, Australia 4103","Canning Vale, Australia 4103","Cartagena, Colombia 4481 Warehouse","Chennai, India 4196","Chicago, IL (65th)","Chicago, IL (65th)","Chongquig, China 4198","Chongquig, China 4198","Cikarang, Indonesia 4165","Cikarang, Indonesia 4165","Clearfield, UT 4851","Crystal Warehouse - MA 4052","Dammam, KSA 4391","Delhi, India 4176","Denver, CO 4997 T2T","Dubai, UAE 4390","Dukinfield, GB 4328","Dukinfield, GB 4328","Duque de Caxias, Brazil 4443","Duque de Caxias, Brazil 4443","Duque de Caxias, Brazil 4443","Edmonton, Albt 4205","El Paso, TX 4973 PBS","Epernon, France 4315 Warehouse","Epping, Australia 4101","Epping, Australia 4101","Epping, Australia 4101","Epping, Australia 4101","Essen, Germany 4362","Ezhou, China 4189","HaiDuong, Vietnam 4199","HaiDuong, Vietnam 4199","HaiDuong, Vietnam 4199","HaiDuong, Vietnam 4199","Halifax, NS (Storage) 4200 Storage","Heist-op-den-Berg, Belgium (DeNeef) 4372","Heist-op-den-Berg, Belgium (DeNeef) 4372","Helsingborg, Sweden 4340 Warehouse","Helsingborg, Sweden 4340 Warehouse","HocMon, Vietnam 4171","HocMon, Vietnam 4171","HocMon, Vietnam 4171","Holden Hill, Australia 4106","Hong Kong (Fanling) 4175","Hong Kong (Fanling) 4175","Houston, TX 4839","Inchon, Korea 4180","Irondale, AL 4619","Jackson, MS 4939 PBS","Jeddah, KSA 4392","Lampa, Chile 4463","Lampa, Chile 4463","Larnaud, France 4317","Lima, Peru 4471 GCP Intl. Sucursal - Cons","Lima, Peru 4471 GCP Intl. Sucursal - Cons","Lithonia, GA 4828","Los Angeles, CA 4848","Luegde, Germany 4361","Mansfield, MA 4915 T2T","Milwaukee, WI 4698 Sales Only","Milwaukee, WI 4698 Sales Only","Minhang, China 4190","Minhang, China 4190","Montreal, QB 4201","Montreal, QB 4201","Mt. Pleasant, TN 4046","North Bergen, NJ 4847","North Bergen, NJ 4847","Olathe, KS 4947 PBS","Panama 4480","Passirana, Italy 4325","Phoenix, AZ 4821 Toller","Phoenix, AZ 4821 Toller","Pompano Beach, FL 4838 Toller","Pompano Beach, FL 4838 Toller","Pompano Beach, FL 4838 Toller","Porirua, New Zealand 4110","Puerto Rico 4909","Quilmes, Argentina 4450","Receife, Brazil 4444 SCC","Receife, Brazil 4444 SCC","Receife, Brazil 4444 SCC","Regents Park Estate, Australia 4102","Regents Park Estate, Australia 4102","Santa Ana, CA 4676","Santiago Tianguistenco, Mexico 4431","Singapore (Jurong) 4150","Singapore (Jurong) 4150","Slough, UK 4322 Office","Somerset, NJ 4663 Warehouse","Sorocaba, Brazil 4440","Spartan, South Africa 4348","Taiwan 4185 - Closed","Toller, Traded Good, Other","Toller, Traded Good, Other","Townsville, Australia 4105","Townsville, Australia 4105","Tuzla, Turkey 4380","Tuzla, Turkey 4380","Valencia, Venzuela 4411","Vancouver, BC 4206 Warehouse","Watermark 4211 Subcontracting","West Chester, OH (Verfi) 4070","West Chester, OH (Verfi) 4070","Widnes, UK 4321","Winnipeg, SK 4203","XiQing, China (Tianjin) 4192","XiQing, China (Tianjin) 4192","XiQing, China (Tianjin) 4192","XiQing, China (Tianjin) 4192","XiQing, China (Tianjin) 4192","Zellwood, FL 4835","Zellwood, FL 4835","ZengCheng, China 4191","ZengCheng, China 4191","Archive, plant not active")

Write-Host $PlantArray.Length

# Created the Search Array
$SearchArray =@("Ajax","Archerfield","Arch","Atsugi","Augusta","Bahia", "Filho", "S.Filho", "Simones","Balakong", "KL", "Malaysia", "Kuala Lumpur","Bangpoo", "Bang","Barcelona","Bellville", "Bell","Bogota","Calhoun","Cambridge","Can", "Canlubang","Canningvale", "Canning Vale", "Cann", "Kewdale", "Kew","Cartagena","Chennai","Bedford Park","Chicago","Chong", "Chongquig","Cikarang", "Cik","Clearfield","Crystal","Dammam","Delhi","Denver","Dubai","Dukin", "Dukinfield","Rio", "Duque", "Caxias","Edmonton","Paso","Epernon","Epping", "Epp", "Fawkner", "Fawk","Essen","Ezhou","Han", "Hanoi", "Hai", "HaiDuong","Halifax","Heist", "Belgium","Helsingborg", "Helsing","HocMon", "HoChiMinh", "Chi","Holden","Hong Kong", "Fanling","Houston","Inchon","Irondale","Jackson","Jeddah","Lampa", "Chile","Larnaud","Lima", "Peru","Lithonia","Angeles","Luegde","Mansfield","Enoree", "Mil","Minhang", "Min","Montreal", "Mont","Pleasant","North", "Bergen","Olathe","Panama","Passirana","Phoen", "Phoenix","Pomp", "Pompano", "Pbeach","Porirua","Puerto","Quilmes","Igarassu", "Recife", "Receife","Reg", "Regents","Santa","Santiago Mex","Singapore", "Jurong","Slough","Somerset","Sorocaba","Spartan","Chung Li","Tolled", "Traded","Town", "Townsville","Tuzla", "Turkey","Valencia","Vancouver","Watermark","West", "Chester","Widnes","Winnipeg","Tianjin", "XiQing", "TanGu", "Tangu", "Tan","Zell", "Zellwood","Zen", "Zeng","San Boi")

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
					$File["FileName"] =$PlantArray[$i] 	 
					$File.update()	
					$clientContext.ExecuteQuery()
					Write-Host "Updating" $($File["FileLeafRef"])
					Write-Host $($File["ID"])
                    "FileName: $($File["FileLeafRef"]), ID:$($File["ID"])" | Add-Content -Path 'C:\Users\amit_dahotre\Desktop\UpdateFileNameLogs.txt'				
                     
				
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
    Break
}
