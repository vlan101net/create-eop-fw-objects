# Function to convert CIDR to dotted decimal subnet mask
Function Convert-RvNetInt64ToIpAddress() 
        { 
            Param 
            ( 
                [int64] 
                $Int64 
            ) 
 
            # Return 
            '{0}.{1}.{2}.{3}' -f ([math]::Truncate($Int64 / 16777216)).ToString(), 
                ([math]::Truncate(($Int64 % 16777216) / 65536)).ToString(), 
                ([math]::Truncate(($Int64 % 65536)/256)).ToString(), 
                ([math]::Truncate($Int64 % 256)).ToString() 
        }

# Function to convert CIDR to dotted decimal subnet mask
Function Convert-RvNetSubnetMaskCidrToClasses 
        { 
            Param 
            ( 
                [int] 
                $SubnetMaskCidr 
            ) 
 
            # Return 
            Convert-RvNetInt64ToIpAddress -Int64 ([convert]::ToInt64(('1' * $SubnetMaskCidr + '0' * (32 - $SubnetMaskCidr)), 2)) 
        }

# Function to download the current list of O365 IP Addresses from Microsoft
Function Get-EOPSubnets
        {
          $addresslist = @()
          $o365json = Invoke-WebRequest -Uri https://endpoints.office.com/endpoints/worldwide?clientrequestid=b10c5ed1-bad1-445f-b386-b919946339a7
          $arrjson = ConvertFrom-Json -InputObject $o365json
          foreach ($objo365 in $arrjson) {
            if ($objo365.ServiceArea -eq "Exchange") {
              foreach ($str_ip in $objo365.ips) {
                $iplist = $str_ip.Split([Environment]::NewLine)
                foreach ($ip in $iplist) {
                  if ($ip -inotmatch ":") {
                    $addresslist += $ip
                  }
                }
              }
            }
          }
          $addresslist = $addresslist | Select -Unique
          return $addresslist
        }

#Function to create commands for address objects and object group for ASA
Function Create-ASACFG($addresslist)
        {
          $counter = 1
          $asafile = "asaeop.txt"
          
          If (Test-Path $asafile) {
        	Remove-Item $asafile
          }

          $asagroup = "object-group network EOP`r`n"
          
          foreach ($address in $addresslist) {
            $name = "EOP-$counter"
            $splitIP = $address.split("/")
            $asaaddress = $splitIP[0]
            $asamask = Convert-RvNetSubnetMaskCidrToClasses $splitIP[1]
            $asacfg = "object network $name`r`nnetwork $asaaddress $asamask"
            $asagroup += "network-object object $name`r`n"
            $asacfg | Out-File $asafile -Append
            $counter++
          }
          
          $asagroup | Out-File $asafile -Append
          write-host "File created: " + $asafile
        }

#Function to create commands for address objects and object group for Fortigate        
Function Create-FORTICFG($addresslist)
        {
          $counter = 1
          $fortifile = "fortieop.txt"
          
          If (Test-Path $fortifile) {
            Remove-Item $fortifile
          }
          
          $fortigroup = "end`r`nconfig firewall addrgrp`r`nedit EOP`r`nset member "
          $fortiaddobj = "config firewall address"
          
          $fortiaddobj | Out-File $fortifile -Append
          
          foreach ($address in $addresslist) {
            $name = "EOP-$counter"
            $splitIP = $address.split("/")
            $fortigroup += "$name "
            $forticfg = "edit $name`r`nset subnet $address`r`nnext"
            $forticfg | Out-File $fortifile -Append
            $counter++           
          }

          $fortigroup += "`r`nnext`r`nend"
          $fortigroup | Out-File $fortifile -Append
          write-host "File created: " + $fortifile
        }

#Function to create commands for address objects and object group for Palo Alto
Function Create-PALOCFG($addresslist)
        {
          $counter = 1
          $palofile = "paloeop.txt"
          
          If (Test-Path $palofile) {
            Remove-Item $palofile
          }

          $palogroup = "set address-group EOP"
          $palogroup | Out-File $palofile -Append

          foreach ($address in $addresslist) {
            $name = "EOP-$counter"
            $splitIP = $address.split("/")
            $paloobjcfg = "set address $name ip-netmask $address`r`n"
            $palogrpcfg = "set address-group EOP static $name`r`n"
            $paloobjcfg | Out-File $palofile -Append
            $palogrpcfg | Out-File $palofile -Append
            $counter++
          }
          write-host "File created: " + $palofile        
        }

#Display user menu to choose output device
Function Show-Menu
        {
          write-host "<===== Choose a Firewall =====>"
          write-host " "
          write-host "1) Cisco ASA"
          write-host "2) Fortigate"
          write-host "3) Palo Alto"
          write-host " "
          write-host "q) Quit"
        }

#Look for input from user and run appropriate routine
Clear-Host
$eopaddresslist = Get-EOPSubnets

do {
  Show-Menu
  $input = Read-Host "Choose a Firewall..."
  
  switch ($input) {
    '1' {
          Clear-Host
          Create-ASACFG($eopaddresslist)
        }
    '2' {
          Clear-Host
          Create-FORTICFG($eopaddresslist)
        }
    '3' {
          Clear-Host
          Create-PALOCFG($eopaddresslist)
        }
  }
}
until ($input -eq 'q')