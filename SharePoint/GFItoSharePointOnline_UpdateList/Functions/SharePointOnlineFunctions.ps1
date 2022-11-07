######################################
## - SharePoint Online Functions  - ##
######################################

##### Add- #####
# SharePoint Online: Add Client side SharePoint Online Management Assemblies
Function Add-SPOAssemblies{
# Sample: Add-SPOAssemblies
   # Requires "SharePoint Online Client Components SDK". Download latest 64 bit version at "https://www.microsoft.com/en-us/download/details.aspx?id=42038"
   # $SPOCSOMPath Is the location for Microsoft.SharePoint.Client.Runtime.dll and Microsoft.SharePoint.Client.dll and it may vary depending on version and installation method
   try{
       $SPOCSOMPath = 'C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI' #"<SPO_Path>"
       $MSSPClient = '\Microsoft.SharePoint.Client.dll'
       $MSSPClientRuntime = '\Microsoft.SharePoint.Client.Runtime.dll'
       Add-Type -Path (Join-Path $SPOCSOMPath $MSSPClient)
       Add-Type -Path (Join-Path $SPOCSOMPath $MSSPClientRuntime)
    }
    Catch{
       Write-Host 'Failed to add .dll files. You will need to install "SharePoint Online Client Components SDK"'
       Write-Host $_.Exception.Message
       Start-Process "https://www.microsoft.com/en-us/download/details.aspx?id=42038"
    }
}

# SharePoint Online: Add Field to list
Function Add-SPOListField{
# Sample: Add-SPOListField  -ListName $ListName -FieldName $FieldName -FieldType $FieldType
   param ($ListName, $FieldName, $FieldType)
   $addNewField = $true
   Try{
      # Get list
      $SPOWeb = $SPOClientContext.Web
      $SPOSite = $SPOClientContext.Site
      $SPOList = $SPOClientContext.Web.Lists.GetByTitle($ListName)
      $SPOListFields = $SPOList.Fields
      $SPOClientContext.Load($SPOWeb)
      $SPOClientContext.Load($SPOSite)
      $SPOClientContext.Load($SPOList)
      $SPOClientContext.Load($SPOListFields)
      $SPOClientContext.ExecuteQuery()

      # Check if list already exists
      if($SPOList -ne $null){
         #Write-Host "List has been found"
         try{
            $SPOListFields | foreach {
               if ($_.InternalName -eq $ListName){
                  $addNewField = $false
               }
            }
         }
         catch{
            Write-Host $_.Exception.Message -ForegroundColor Red
         }
         if($addNewField -eq $true){
            # Add Field to list
            $SPOList = $SPOClientContext.Web.Lists.GetByTitle($ListName)
            $FieldXml = "<Field Type='$FieldType' DisplayName='$FieldName' />"
            $Option = [Microsoft.SharePoint.Client.AddFieldOptions]::AddFieldToDefaultView
            $SPOList.Fields.AddFieldAsXml($fieldxml,$true,$option) | Out-Null # Hiding itemcreation details.
            $SPOClientContext.Load($SPOList)
            $SPOClientContext.ExecuteQuery()
            Write-Host "Created: $FieldName in $ListName" -ForegroundColor Cyan
         }
         else{
            # If the field with the same name already exists
            write-host "Field '$FieldName' Already Exists in the List" -ForegroundColor Yellow
         }
      }
      else{
         # If the list cannot be found
         write-host "List '$ListName' doesn't exists!" -ForegroundColor Yellow
      }
   }
   Catch{
      Write-Host $_.Exception.Message -ForegroundColor Red
   }
}

# SharePoint Online: Dynamicly add fieldvalues to list 
Function Add-SPOListItem{
# Sample: Add-SPOListItem -ListName $ListName -Values @{
#            "Title" = $workstation.name.InnerText;
#            "Client" = $client.name.InnerText;
#            "ip_public" = $workstation.external_ip.InnerText;
#            "ip_private" = $workstation.ip.InnerText          
#         }
   param ($ListName, $Values)
   Try{
      # Get list
      $List = Get-SPOListItem -List $ListName

      if($Values.Count -gt 0){
         $ListItemCreationInformation = New-Object Microsoft.SharePoint.Client.ListItemCreationInformation
         $NewListItem = $List.AddItem($ListItemCreationInformation)

         # Dynamicly add values to fields
         $Values.Keys | foreach{
            $FieldName = ('{0}' -f $_,$Values[$_]).ToString()
            $Value = ('{1}' -f $_,$Values[$_]).ToString()
            $NewListItem[$FieldName] = $Value
            $NewListItem.Update()
            Write-Host $FieldName $Value -ForegroundColor Green
         }

         # Execute changes
         $SPOClientContext.ExecuteQuery()
      }
      else{
         Write-Host "Field item empty" -ForegroundColor Yellow
      }
   }
   Catch{
      Write-Host $_.Exception.Message -ForegroundColor Red
   }
}

##### Connect- #####
# Connect to SharePoint Online Site
Function Connect-SPOnline{
# Sample: $SPOClientContext = Connect-SPOnline -SPOSiteUrl "$SPOSiteUrl" -SPOUserName "$SPOUserName" -SPOPassword "$SPOPassword"
   # Important: This function returns "$SPOClientContext"
   param($SPOSiteUrl, $SPOUserName, $SPOPassword)
   try{
      # Import important stuff
      ## For this to work you may need to download and install "SharePoint Online Management Shell"
      [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client") | Out-Null
      [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client.Runtime") | Out-Null
      Import-Module Microsoft.Online.SharePoint.PowerShell -WarningAction SilentlyContinue
      
      #Adding the Client side SharePoint Online Management Assemblies
      Add-SPOAssemblies

      # Protect Password
      $SPOPassword = $SPOPassword | ConvertTo-SecureString -AsPlainText -Force
      $SPOCredentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($SPOUserName, $SPOPassword)

      # Bind to site collection
      $SPOClientContext = New-Object Microsoft.SharePoint.Client.ClientContext($SPOSiteUrl)
      $SPOClientContext.Credentials = $SPOCredentials
      $SPOClientContext.ExecuteQuery()
      Write-Host "Connected to $SPOSiteUrl!"
      Return $SPOClientContext
   }
   Catch{
      Write-Host "Could not connect to $SPOSiteUrl : Check login ID, password and connection" -ForegroundColor Magenta
      Write-Host $_.Exception.Message

   }
}

##### Create- #####
# SharePoint Online: Create new list
function Create-SPOList{
# Sample: Create-SPOList -SPOSiteUrl $SPOSiteUrl -SPOListName $SPOListName
   param ($SPOSiteUrl, $ListName)
   try{        
       #SPO Client Object Model Context
       Write-Host "Creating List $ListName in $SPOSiteUrl" -ForegroundColor Green
       $SPOWeb=$SPOClientContext.Web 
       $SPOListCreationInformation = New-Object Microsoft.SharePoint.Client.ListCreationInformation 
       $SPOListCreationInformation.Title = $ListName 
        
       #https://msdn.microsoft.com/EN-US/library/office/microsoft.sharepoint.client.listtemplatetype.aspx 
       $SPOListCreationInformation.TemplateType = [int][Microsoft.SharePoint.Client.ListTemplatetype]::GenericList 
       $SPOList=$SPOWeb.Lists.Add($SPOListCreationInformation)

       # Make changes
       $SPOClientContext.ExecuteQuery()
       Write-Host "List $ListName created in $SPOSiteUrl" -ForegroundColor Green
   }
   catch [System.Exception] 
   { 
      Write-Host -ForegroundColor Red $_.Exception.ToString()
   }
}

##### Enumerate- #####
Function Enumerate-SPOList(){
# Sample: $EnumList = Enumerate-SPOList -ListName $ListName -SearchField $SearchField
   param ($ListName, $SearchField)
   Try{
      Write-Host "Attemting to enumerate $ListName" -ForegroundColor Cyan

      # Get list
      $SPOWeb = $SPOClientContext.Web
      $SPOSite = $SPOClientContext.Site

      # Filter and get the List Items using CAML
      $SPOList = $SPOClientContext.Web.Lists.GetByTitle($ListName)
      $SPOListFields = $SPOList.Fields
      $SPOClientContext.Load($SPOWeb)
      $SPOClientContext.Load($SPOSite)
      $SPOClientContext.Load($SPOList)
      $SPOClientContext.Load($SPOListFields)
      $SPOClientContext.ExecuteQuery()

      # Make EnumList
      $ItemCount = $SPOList.ItemCount
      Write-Host "Making EnumList for $ListName" -ForegroundColor Cyan
      $NewEnumList = @()
      for ($i = 1; $i -le $ItemCount; $i++){
         # Get List Items by ID
         try{
            $CurrentItem = $SPOList.GetItemById($i)
            $SPOClientContext.Load($CurrentItem)
            $SPOClientContext.ExecuteQuery()
            $NewEnumList += $CurrentItem[$SearchField] +","+ $CurrentItem["ID"]
            Write-Host $CurrentItem[$SearchField] "ID" $CurrentItem["ID"]
         }
         Catch{
            Write-Host "Item ID: $i Not found. Likely deleted!" -BackgroundColor Red -ForegroundColor White
            # Adding to counter for each missing item until all available items are found
            $ItemCount += 1
         }
      }
      # Returns an array
      Write-Host "Enumeration completed: $ListName" -ForegroundColor Green
      Return $NewEnumList
   }
   Catch{
      Write-Host "Error" -ForegroundColor Yellow
      Write-Host $_.Exception.Message -ForegroundColor Red
   }
}

##### Get- #####
Function Get-SPOListItem{
# Sample: $List = Get-SPOListItem -List "list_name"
   param($List)
   try{
      # Get list
      $SPOWeb = $SPOClientContext.Web
      $SPOSite = $SPOClientContext.Site
      $SPOList = $SPOClientContext.Web.Lists.GetByTitle($List)
      $SPOListFields = $SPOList.Fields
      $SPOClientContext.Load($SPOWeb)
      $SPOClientContext.Load($SPOSite)
      $SPOClientContext.Load($SPOList)
      $SPOClientContext.Load($SPOListFields)
      $SPOClientContext.ExecuteQuery()

      Return $SPOList
   }
   Catch{
      Write-Host $_.Exception.Message -ForegroundColor Red
   }
}

##### Invoke- #####
# Execute query
Function Invoke-SPOExecuteQuery(){
# Sample: Invoke-SPOExecuteQuery
   try{
      # Execute changes
      $SPOClientContext.ExecuteQuery()
      # Clean up
      $SPOClientContext.Dispose()
   }
   Catch{
      Write-Host $_.Exception.Message
   }
}

##### Remove-  #####
# SharePoint Online: Remove list
Function Remove-SPOList{
# Sample: Remove-SPOList "name of list to be removed"
   param ($ListToBeRemoved)
   try{
      $Listist = $SPOClientContext.Web.Lists.GetByTitle($ListToBeRemoved)
      $SPOClientContext.Load($List)
      $Fields = $List.Fields
      $List.DeleteObject()
      $SPOClientContext.ExecuteQuery()
   }
   catch{
      Write-Host -ForegroundColor Red $_.Exception.ToString()
   }
}

##### Test-    #####
# Test if List exists on a Site
Function Test-SPOListExistence{
# Sample: Test-SPOListExistence -ListName $ListName
   param ($ListName)
   $ListExists = $false

   # Get lists
   $SPOWeb = $SPOClientContext.Web
   $SPOSite = $SPOClientContext.Site
   $SPOLists = $SPOClientContext.Web.Lists
   $SPOClientContext.Load($SPOWeb)
   $SPOClientContext.Load($SPOSite)
   $SPOClientContext.Load($SPOLists)
   $SPOClientContext.ExecuteQuery()
   $ListLists = $SPOLists | Select -Property Title

   # Compare $ListName with Lists
   for($i = 0; $i -lt $ListLists.Count; $i++){
      $currentItem = $listLists[$i] | Select -Property Title
      if($currentItem.Title -eq $ListName){
         Write-Host $currentItem.Title " Exists!"
         $ListExists = $true
      }
   }

   # Create  new list?
   if($ListExists -eq $false){
      if($Autocreate -eq $true){
         # Creates a new List without asking permition
         Create-SPOList -SPOSiteUrl $SPOSiteUrl -ListName $ListName
      }
      if($AutoCreate -eq $false){
         # Asks for permition before creating a new List
         $ans = Read-Host "$ListName does not exist on $SPOSiteURL. Do you want to create it? (Y/N)"
         if($ans -eq "y" -or "Y"){
            Create-SPOList -SPOSiteUrl $SPOSiteUrl -ListName $ListName
         }
      }
   }
}

# SharePoint Online: Test if List or fields are missing
Function Test-SPOListAndFieldExistence{
# Sample: Test-SPOListAndFieldExistence -ListName $ListName -Values @{
#            "Title" = $workstation.name.InnerText;
#            "Client" = $client.name.InnerText;
#            "ip_publik" = $workstation.external_ip.InnerText;
#            "ip_privat" = $workstation.ip.InnerText          
#         }
   param ($ListName, $Values)
   Try{
      # Check if the list exists
      Test-SPOListExistence -ListName $ListName
      # Get list
      $List = Get-SPOListItem -List $ListName

      if($Values.Count -gt 0){
         $ListItemCreationInformation = New-Object Microsoft.SharePoint.Client.ListItemCreationInformation
         $NewListItem = $List.AddItem($ListItemCreationInformation)

         # Dynamicly add values to fields
         $Values.Keys | foreach{
            $FieldName = ('{0}' -f $_,$Values[$_]).ToString()

            # Test if field exists in the list
            Test-SPOListFieldExistence -ListName $ListName -FieldName $FieldName
         }
      }
      else{
         Write-Host "Field item empty" -ForegroundColor Magenta
      }
   }
   Catch{
      Write-Host $_.Exception.Message -ForegroundColor Red
   }
}

# Check if a field exists in a list
Function Test-SPOListFieldExistence{
# Sample Test-SPOListFieldExistence -ListName $ListName -FieldName $FieldName
   param ($ListName, $FieldName)
   $FieldExists = $false
   $SPOWeb = $SPOClientContext.Web
   $SPOSite = $SPOClientContext.Site
   $SPOList = $SPOClientContext.Web.Lists.GetByTitle($ListName)
   $SPOListFields = $SPOList.Fields
   $SPOClientContext.Load($SPOWeb)
   $SPOClientContext.Load($SPOSite)
   $SPOClientContext.Load($SPOList)
   $SPOClientContext.Load($SPOListFields)
   $SPOClientContext.ExecuteQuery()
   $SPOListFields | foreach{
      if($_.InternalName -eq $FieldName){
         $FieldExists = $true
      }
   }
   # Create new field?
   if($FieldExists -eq $false){
      if($Autocreate -eq $true){
         # Just creates new fields without asking!
         Add-SPOListField  -ListName $ListName -FieldName $FieldName -FieldType "Text"
      }
      if($Autocreate -eq $false){
      # Asks for each new field before creating them
         $ans = Read-Host "$FieldName does not exist in $ListName. Do you want to create it? (Y/N)"
         if($ans -eq "y" -or "Y"){
            Add-SPOListField  -ListName $ListName -FieldName $FieldName -FieldType "Text"
         }
      }

   }
   else{
      Write-Host "$FieldName found in $ListName."
   }
}

##### Update-  #####
# SharePoint Online: Update values of fields in lists - Needs fixing
Function Update-SPOListFielItems{
# Sample: Update-SPOListFielItems -ListName $ListName -SearchField $SearchField -EnumList $EnumList -Values $Values 
   param ($ListName, $SearchField, $EnumList, $Values)
   try{
      # Get list
      $SPOWeb = $SPOClientContext.Web
      $SPOSite = $SPOClientContext.Site
      # Filter and get the List Items using CAML
      $SPOList = $SPOClientContext.Web.Lists.GetByTitle($ListName)
      $SPOListFields = $SPOList.Fields
      $SPOClientContext.Load($SPOWeb)
      $SPOClientContext.Load($SPOSite)
      $SPOClientContext.Load($SPOList)
      $SPOClientContext.Load($SPOListFields)
      $SPOClientContext.ExecuteQuery()

      # Resetting the following variables
      $SearchValue = ""
      $MissingItem = $true

      # Get current SearchValue
      Try{
         if(($Values.Count -gt 0) -and ($EnumList.Count -gt 0)){
            $Values.Keys | foreach{
               $FieldName = ('{0}' -f $_,$Values[$_]).ToString()
               $Value = ('{1}' -f $_,$Values[$_]).ToString()
               if($FieldName -eq $SearchField){
                  $SearchValue = $Value
               }
            }
            Write-Host "Looking for UID: $SearchValue" -ForegroundColor Cyan
            for([Int64]$i = 0; ($i -ge 0) -and ($i -le $EnumList.Count) ; $i++){
               $newItem = SplitLeft $EnumList[$i]
               if($newItem -eq $SearchValue){
                  # update item
                  Write-Host "Updating Item" -ForegroundColor Green

                  # Declaring item not missing! Avoids creating duplicates
                  $MissingItem = $false
                  $ItemID = SplitRight $EnumList[$i]
                  $ItemID = [Int]$ItemID
                  Write-Host "UID: $SearchValue - SPOListID: $ItemID" -ForegroundColor Green

                  # Get List Items by ID
                  $ListItem = $SPOList.GetItemById($ItemID)

                  # Dynamicly add values to fields
                  $Values.Keys | foreach{
                     $FieldName = ('{0}' -f $_,$Values[$_]).ToString()
                     $Value = ('{1}' -f $_,$Values[$_]).ToString()
                     $ListItem[$FieldName] = $Value
                     $ListItem.Update()
                     Write-Host $FieldName $Value -ForegroundColor White
                  }
                  # Make changes
                  $SPOClientContext.ExecuteQuery()
                  # Exit loop
                  $i = -10
               }
            }
         }
      }
      Catch{
         Write-Host "Item not found in $ListName" -ForegroundColor Yellow
      }
      if($MissingItem -eq $true){
            Write-Host "Creating new Item" -ForegroundColor Yellow
            Add-SPOListItem -ListName $ListName -Values $Values
         }
      else{
         Write-Host "No new Item Created" -ForegroundColor Yellow
      }
   }
   Catch{
      Write-Host "Error" -ForegroundColor Yellow
      Write-Host $_.Exception.Message -ForegroundColor Red
   }
}

#####> End of File <#####
