Function ConvertTo-BinaryIP {
  <#
    .Synopsis
      Converts a Decimal IP address into a binary format.
    .Description
      ConvertTo-BinaryIP uses System.Convert to switch between decimal and binary format. The output from this function is dotted binary.
    .Parameter IPAddress
      An IP Address to convert.
  #>
 
  [CmdLetBinding()]
  Param(
    [Parameter(Mandatory = $True, Position = 0, ValueFromPipeline = $True)]
    [Net.IPAddress]$IPAddress
  )
 
  Process {  
    Return [String]::Join('.', $( $IPAddress.GetAddressBytes() |
      ForEach-Object { [Convert]::ToString($_, 2).PadLeft(8, '0') } ))
  }
}

Function ConvertTo-DecimalIP {
  <#
    .Synopsis
      Converts a Decimal IP address into a 32-bit unsigned integer.
    .Description
      ConvertTo-DecimalIP takes a decimal IP, uses a shift-like operation on each octet and returns a single UInt32 value.
    .Parameter IPAddress
      An IP Address to convert.
  #>
   
  [CmdLetBinding()]
  Param(
    [Parameter(Mandatory = $True, Position = 0, ValueFromPipeline = $True)]
    [Net.IPAddress]$IPAddress
  )
 
  Process {
    $i = 3; $DecimalIP = 0;
    $IPAddress.GetAddressBytes() | ForEach-Object { $DecimalIP += $_ * [Math]::Pow(256, $i); $i-- }
 
    Return [UInt32]$DecimalIP
  }
}

Function ConvertTo-DottedDecimalIP {
  <#
    .Synopsis
      Returns a dotted decimal IP address from either an unsigned 32-bit integer or a dotted binary string.
    .Description
      ConvertTo-DottedDecimalIP uses a regular expression match on the input string to convert to an IP address.
    .Parameter IPAddress
      A string representation of an IP address from either UInt32 or dotted binary.
  #>
 
  [CmdLetBinding()]
  Param(
    [Parameter(Mandatory = $True, Position = 0, ValueFromPipeline = $True)]
    [String]$IPAddress
  )
   
  Process {
    Switch -RegEx ($IPAddress) {
      "([01]{8}\.){3}[01]{8}" {
        Return [String]::Join('.', $( $IPAddress.Split('.') | ForEach-Object { [Convert]::ToUInt32($_, 2) } ))
      }
      "\d" {
        $IPAddress = [UInt32]$IPAddress
        $DottedIP = $( For ($i = 3; $i -gt -1; $i--) {
          $Remainder = $IPAddress % [Math]::Pow(256, $i)
          ($IPAddress - $Remainder) / [Math]::Pow(256, $i)
          $IPAddress = $Remainder
         } )
        
        Return [String]::Join('.', $DottedIP)
      }
      default {
        Write-Error "Cannot convert this format"
      }
    }
  }
}

Function ConvertTo-MaskLength {
  <#
    .Synopsis
      Returns the length of a subnet mask.
    .Description
      ConvertTo-MaskLength accepts any IPv4 address as input, however the output value 
      only makes sense when using a subnet mask.
    .Parameter SubnetMask
      A subnet mask to convert into length
  #>
 
  [CmdLetBinding()]
  Param(
    [Parameter(Mandatory = $True, Position = 0, ValueFromPipeline = $True)]
    [Alias("Mask")]
    [Net.IPAddress]$SubnetMask
  )
 
  Process {
    $Bits = "$( $SubnetMask.GetAddressBytes() | ForEach-Object { [Convert]::ToString($_, 2) } )" -Replace '[\s0]'
 
    Return $Bits.Length
  }
}

Function ConvertTo-Mask {
  <#
    .Synopsis
      Returns a dotted decimal subnet mask from a mask length.
    .Description
      ConvertTo-Mask returns a subnet mask in dotted decimal format from an integer value ranging 
      between 0 and 32. ConvertTo-Mask first creates a binary string from the length, converts 
      that to an unsigned 32-bit integer then calls ConvertTo-DottedDecimalIP to complete the operation.
    .Parameter MaskLength
      The number of bits which must be masked.
  #>
   
  [CmdLetBinding()]
  Param(
    [Parameter(Mandatory = $True, Position = 0, ValueFromPipeline = $True)]
    [Alias("Length")]
    [ValidateRange(0, 32)]
    $MaskLength
  )
   
  Process {
    Return ConvertTo-DottedDecimalIP ([Convert]::ToUInt32($(("1" * $MaskLength).PadRight(32, "0")), 2))
  }
}

Function Get-NetworkAddress {
  <#
    .Synopsis
      Takes an IP address and subnet mask then calculates the network address for the range.
    .Description
      Get-NetworkAddress returns the network address for a subnet by performing a bitwise AND 
      operation against the decimal forms of the IP address and subnet mask. Get-NetworkAddress
      expects both the IP address and subnet mask in dotted decimal format.
    .Parameter IPAddress
      Any IP address within the network range.
    .Parameter SubnetMask
      The subnet mask for the network.
  #>
   
  [CmdLetBinding()]
  Param(
    [Parameter(Mandatory = $True, Position = 0, ValueFromPipeline = $True)]
    [Net.IPAddress]$IPAddress,
     
    [Parameter(Mandatory = $True, Position = 1)]
    [Alias("Mask")]
    [Net.IPAddress]$SubnetMask
  )
 
  Process {
    Return ConvertTo-DottedDecimalIP ((ConvertTo-DecimalIP $IPAddress) -BAnd (ConvertTo-DecimalIP $SubnetMask))
  }
}

Function Get-BroadcastAddress {
  <#
    .Synopsis
      Takes an IP address and subnet mask then calculates the broadcast address for the range.
    .Description
      Get-BroadcastAddress returns the broadcast address for a subnet by performing a bitwise AND 
      operation against the decimal forms of the IP address and inverted subnet mask. 
      Get-BroadcastAddress expects both the IP address and subnet mask in dotted decimal format.
    .Parameter IPAddress
      Any IP address within the network range.
    .Parameter SubnetMask
      The subnet mask for the network.
  #>
   
  [CmdLetBinding()]
  Param(
    [Parameter(Mandatory = $True, Position = 0, ValueFromPipeline = $True)]
    [Net.IPAddress]$IPAddress, 
     
    [Parameter(Mandatory = $True, Position = 1)]
    [Alias("Mask")]
    [Net.IPAddress]$SubnetMask
  )
 
  Process {
    Return ConvertTo-DottedDecimalIP $((ConvertTo-DecimalIP $IPAddress) -BOr `
      ((-BNot (ConvertTo-DecimalIP $SubnetMask)) -BAnd [UInt32]::MaxValue))
  }
}

Function Get-NetworkSummary ( [String]$IP, [String]$Mask ) {
  If ($IP.Contains("/"))
  {
    $Temp = $IP.Split("/")
    $IP = $Temp[0]
    $Mask = $Temp[1]
  }
 
  If (!$Mask.Contains("."))
  {
    $Mask = ConvertTo-Mask $Mask
  }
 
  $DecimalIP = ConvertTo-DecimalIP $IP
  $DecimalMask = ConvertTo-DecimalIP $Mask
   
  $Network = $DecimalIP -BAnd $DecimalMask
  $Broadcast = $DecimalIP -BOr
    ((-BNot $DecimalMask) -BAnd [UInt32]::MaxValue)
  $NetworkAddress = ConvertTo-DottedDecimalIP $Network
  $RangeStart = ConvertTo-DottedDecimalIP ($Network + 1)
  $RangeEnd = ConvertTo-DottedDecimalIP ($Broadcast - 1)
  $BroadcastAddress = ConvertTo-DottedDecimalIP $Broadcast
  $MaskLength = ConvertTo-MaskLength $Mask
   
  $BinaryIP = ConvertTo-BinaryIP $IP; $Private = $False
  Switch -RegEx ($BinaryIP)
  {
    "^1111"  { $Class = "E"; $SubnetBitMap = "1111" }
    "^1110"  { $Class = "D"; $SubnetBitMap = "1110" }
    "^110"   { 
      $Class = "C"
      If ($BinaryIP -Match "^11000000.10101000") { $Private = $True } }
    "^10"    { 
      $Class = "B"
      If ($BinaryIP -Match "^10101100.0001") { $Private = $True } }
    "^0"     { 
      $Class = "A"
      If ($BinaryIP -Match "^00001010") { $Private = $True } }
   }   
    
  $NetInfo = New-Object Object
  Add-Member NoteProperty "Network" -Input $NetInfo -Value $NetworkAddress
  Add-Member NoteProperty "Broadcast" -Input $NetInfo -Value $BroadcastAddress
  Add-Member NoteProperty "Range" -Input $NetInfo `
    -Value "$RangeStart - $RangeEnd"
  Add-Member NoteProperty "Mask" -Input $NetInfo -Value $Mask
  Add-Member NoteProperty "MaskLength" -Input $NetInfo -Value $MaskLength
  Add-Member NoteProperty "Hosts" -Input $NetInfo `
    -Value $($Broadcast - $Network - 1)
  Add-Member NoteProperty "Class" -Input $NetInfo -Value $Class
  Add-Member NoteProperty "IsPrivate" -Input $NetInfo -Value $Private
   
  Return $NetInfo
}

Function Get-NetworkRange( [String]$IP, [String]$Mask ) {
  If ($IP.Contains("/"))
  {
    $Temp = $IP.Split("/")
    $IP = $Temp[0]
    $Mask = $Temp[1]
  }
 
  If (!$Mask.Contains("."))
  {
    $Mask = ConvertTo-Mask $Mask
  }
 
  $DecimalIP = ConvertTo-DecimalIP $IP
  $DecimalMask = ConvertTo-DecimalIP $Mask
   
  $Network = $DecimalIP -BAnd $DecimalMask
  $Broadcast = $DecimalIP -BOr ((-BNot $DecimalMask) -BAnd [UInt32]::MaxValue)
 
  For ($i = $($Network + 1); $i -lt $Broadcast; $i++) {
    ConvertTo-DottedDecimalIP $i
  }
}

Function Add-DHCPCustomScope {
  <#
    .Synopsis
      ???
    .Description
      ?
    .EXAMPLE
    Add-DHCPCustomScope -i "10.123.151.0" -su "24" -sc "I" -dh "dc-07" -Add "info-terminal" -Venue "SLV"
    .EXAMPLE
  #>
   
  [CmdLetBinding()]
  Param(
    [Parameter(Mandatory = $True, Position = 0, ValueFromPipeline = $True)]
    [String]$IPAddress, 
     
    [Parameter(Mandatory = $True, Position = 1, ValueFromPipeline = $True)]
    [String]$SubnetMask,

    [Parameter(Mandatory = $True, Position = 2, ValueFromPipeline = $True)]
    [String]$ScopeType,

    [Parameter(Mandatory = $True, Position = 3, ValueFromPipeline = $True)]
    [String]$DHCP,

    [Parameter(Mandatory = $True, Position = 3, ValueFromPipeline = $True)]
    [String]$Venue,

    [Parameter(Mandatory = $False, Position = 3, ValueFromPipeline = $True)]
    [String]$Add
  )

$ScopeInfo = Get-NetworkSummary $IPAddress $SubnetMask
$ScopeInfo
$FirstIPAddress = [regex]::Replace($ScopeInfo.Range,"\s.*","0")
$FirstIPAddress
$Gateway = [regex]::Replace($ScopeInfo.Range,"\s.*","")
$Gateway
$LastIPAddress = [regex]::Replace($ScopeInfo.Range,"^\S*\s-\s","")
$LastIPAddress
$DHCP = "\\" + $DHCP
&netsh dhcp server $DHCP add scope $IPAddress $ScopeInfo.Mask Temp Temp | Out-Null
&netsh dhcp server $DHCP scope $IPAddress add iprange $FirstIPAddress $LastIPAddress | Out-Null
&netsh dhcp server $DHCP scope $IPAddress set optionvalue 003 IPADDRESS $Gateway | Out-Null

 Switch -regex ($ScopeType) {
 "A" {
    $NetArea = "Admin"
    &netsh dhcp server $DHCP scope $IPAddress set optionvalue 252 STRING  "http://172.16.22.150/wpad.dat" | Out-Null
    }
 "I" {
    $NetArea = "Games"
    $NetType = $Add
    &netsh dhcp server $DHCP scope $IPAddress set optionvalue 242 STRING  "L2Q=1,L2QVLAN=48" | Out-Null
    &netsh dhcp server $DHCP scope $IPAddress set optionvalue 252 STRING  "http://172.16.22.150/wpad.dat" | Out-Null
    }
 "G" {
    $NetArea = "Guest"
    &netsh dhcp server $DHCP scope $IPAddress set optionvalue 006 IPADDRESS 8.8.8.8 8.8.4.4 | Out-Null
    }
 "M" {
    $NetArea = "Mng"
    }
 "V" {
    $NetType = "Voice"
    &netsh dhcp server $DHCP scope $IPAddress set optionvalue 242 STRING  "L2Q=1,L2QVLAN=48,MCIPADD=10.29.33.60,10.3.0.2,HTTPSRVR=172.16.22.31,HTTPDIR= /http96xx/" | Out-Null
    }
 "W" {
    $NetType = "WiFi"
    &netsh dhcp server $DHCP scope $IPAddress set optionvalue 043 BYTE  "" | Out-Null
    }
 }
 $NetArea
 $NetType
 $ScopeName = "AVAYA_" + $NetArea + "_" + $NetType + "_" + $Venue
 $ScopeName
 &netsh dhcp server $DHCP scope $IPAddress set comment $ScopeName | Out-Null
 &netsh dhcp server $DHCP scope $IPAddress set name $ScopeName | Out-Null
 &netsh dhcp server $DHCP scope $IPAddress set state 1 | Out-Null
}