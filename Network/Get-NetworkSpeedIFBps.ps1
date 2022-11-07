Clear-Host
$NotGbps = $true 
While ($NotGbps){
   $NicInfo = Get-NetAdapter | Select-Object InterfaceDescripion, Name, LinkSpeed
   $TextColor = "White"
   for($i=0;$i -lt $NicInfo.Count; $i++){
      if($NicInfo[$i].LinkSpeed -eq "1 Gbps"){
         #$NotGbps = $false
         $TextColor = "Green"
      }
      Write-Host $NicInfo[$i].Name $NicInfo[$i].LinkSpeed -ForegroundColor $TextColor
   }
   Write-Output $NicInfo -ForegroundColor $TextColor
   Start-Sleep -Seconds 5
}

$console.foregroundcolor = "White"
Write-Output $NicInfo

$console = $Host.UI.RawUI