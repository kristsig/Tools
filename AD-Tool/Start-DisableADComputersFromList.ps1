# Edit to where you like to keep the textfile
$PathToComputerList = 'C:\Users\'+ $env:UserName +'\Desktop\DisableList.txt'
# Depending on organisation naming rules
$ComputerName_MinLength = 6

Import-Module ActiveDirectory
Clear-Host

Get-Content -Path $PathToComputerList | ForEach-Object{
   if($_.length -ge $ComputerName_MinLength){
      try{
          Set-ADComputer -Identity $_ -Enabled $false
          Write-Host "Disabel: $_"
      }
      catch{
         Write-Host $_
      }
   }
   else{
      Write-Host $_ "to short"
   }
}