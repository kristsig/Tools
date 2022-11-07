# Prevent memoryleakage
$Host.Runspace.ThreadOptions = “ReuseThread”

# Timer to evaluate the time it takes to run the script
$StopWatch = [system.diagnostics.stopwatch]::StartNew()

# Include .ps1 files
$ScriptDirectory = (Resolve-Path .\).Path
Try{
   .("$ScriptDirectory\Settings.ps1")
   .("$ScriptDirectory\Var\Credentials.ps1")
   .("$ScriptDirectory\Functions\SplitString.ps1")
   .("$ScriptDirectory\Functions\SharePointOnlineFunctions.ps1")
   .("$ScriptDirectory\Functions\EpochDate.ps1")
   Clear-Host
}
catch {
    Write-Host "Error while loading supporting PowerShell Scripts" -ForegroundColor Magenta
    Write-Host $_.Exception.Message
}

##### Create NewArray
$Values = ItemArray

# Connect
$SPOClientContext = Connect-SPOnline -SPOSiteUrl "$SPOSiteUrl" -SPOUserName "$SPOUserName" -SPOPassword "$SPOPassword"

# Checks if List exists on the site and checks if Fields exist in the List
Test-SPOListAndFieldExistence -ListName $ListName -Values $Values

# Enumerate existing List items
$EnumList = New-Object System.Collections.ArrayList
$EnumList = Enumerate-SPOList -ListName $ListName -SearchField $SearchField


###############

#Adress till clients
$clientsUrl = New-Object System.Xml.XmlDocument
$clientsUrl.Load("https://yourURL")
$clients = $clientsUrl.result.items.client

#Loopa igenom alla clients
foreach ($client in $clients){
   $clientid = $client.clientid
   $clientname = $client.name.InnerText

   #debug
   #Write-Host $clientid
   Write-Host $clientname

   #Adress till sites
   $sitesUrl = New-Object System.Xml.XmlDocument
   $sitesUrl.Load("https://yourURL")
   $sites = $sitesUrl.result.items.site

   foreach ($site in $sites){
      $sitename = $site.name.InnerText
      $siteid = $site.siteid

      #debug
      Write-Host $sitename
      #Write-Host $siteid

      #Adress till workstations
      $workstationsUrl = New-Object System.Xml.XmlDocument
      $workstationsUrl.Load("https://YourURL")
      $workstations = $workstationsURL.result.items.workstation

      #Loopa igenom workstations och lägg till i SP-listan
      foreach ($workstation in $workstations){
         $os = $workstation.os.InnerText

         if ($os){
         $os = $os.replace("Microsoft Windows","Win")
         $os = $os.replace("Professional","Pro")
         $os = $os.replace("Microsoft? Windows Vista?","Vista")
          }

          $manufacturer = $workstation.manufacturer.InnerText

          if ($manufacturer) {
             $manufacturer =  $manufacturer.replace("Hewlett-Packard","HP")
             $manufacturer = $manufacturer.replace("ASUSTeK COMPUTER INC.","Asus")
          }

          $model = $workstation.model.InnerText

          if ($model) {
             $model = $model.replace("HP Compaq ","")
             $model = $model.replace("HP ","")
             $ram = $workstation.total_memory
             $ram = $ram / 1024 / 1024 / 1024
          }
          
          ##### Recreate ItemArray for each Item
          $Values = ItemArray
          Write-Host " - "$clientname $workstation.name.InnerText $workstation.external_ip.InnerText $workstation.ip.InnerText " - " -ForegroundColor Green
          Update-SPOListFielItems -ListName $ListName -SearchField $SearchField -EnumList $EnumList -Values $Values
      #workstations loop stopp
      }
   #sites loop stopp
   }
#clients loop stopp
}

#End#>
$StopWatch.Stop()
Write-Host 'Exit Ƹ̵̡Ӝ̵̨̄Ʒ'
Write-Host $StopWatch.Elapsed
exit
