##############
#  Settings  #
##############

# State the name for the list
$ListName = "Datorer"

# State the name of the "column" where "unique identifier" is contained for each item.
$SearchField = "UID"

# For creating Fields for the List
# Sample:
# $Values = @{
#    "Column name" = $variablecontainingusefulinfo
# }
Function ItemArray{
   $Values = @{
             "Title" = $workstation.name.InnerText;
             "Kund" = $clientname;
             "IP_Publik" = $workstation.external_ip.InnerText;
             "IP_Privat" = $workstation.ip.InnerText;
             "SN" = $workstation.device_serial.InnerText;
             "Site" = $sitename;
             "Modell" = $model;
             "UID" = $workstation.guid.InnerText;
             "Beskrivning" = $workstation.description.InnerText;
             "OS" = $os;
             "Domain" = $workstation.domain.InnerText;
             "User" = $workstation.user.InnerText;
             "Tillverkare" = $manufacturer;
             "Carepack" = $workstation.carepack.InnerText;
             "RAM" = $ram;
             #"CPU" = $workstation.processor_count;
             #"HD" = $workstation.dsc_status;
             "Installations_Datum" = $workstation.install_date;
             "Scanned" = $workstation.last_scan_time;
             "Boot_time" = (Get-EpochDate $workstation.last_boot_time)
          }
   return $Values
}

# Automaticly create new list with all fields if $true, Asks before creating new List and Fields if $false
$Autocreate = $true