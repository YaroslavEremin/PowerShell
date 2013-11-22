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
    .Parameter IPAddress
      ?
    .Parameter SubnetMask
      ?
  #>
   
  [CmdLetBinding()]
  Param(
    [Parameter(Mandatory = $True, Position = 0, ValueFromPipeline = $True)]
    [Net.IPAddress]$IPAddress, 
     
    [Parameter(Mandatory = $True, Position = 1, ValueFromPipeline = $True)]
    [Net.IPAddress]$SubnetMask,

    [Parameter(Mandatory = $True, Position = 2, ValueFromPipeline = $True)]
    [Net.IPAddress]$ScopeType,

    [Parameter(Mandatory = $True, Position = 3, ValueFromPipeline = $True)]
    [Net.IPAddress]$Venue
  )

$ScopeInfo = Get-NetworkSummary $IPAddress $SubnetMask
$ScopeName = "AVAYA_" + $NetArea + "_" + $NetType + "_" + $Venue
$FirstIPAddress = [regex]::Replace($ScopeInfo.Range,"\s.*","0")
$Gateway = [regex]::Replace($ScopeInfo.Range,"\s.*","")
$LastIPAddress = [regex]::Replace($ScopeInfo.Range,"^\S*\s-\s","")

&netsh dhcp server $DHCP add scope $IPAddress $ScopeInfo.Mask Temp Temp
&netsh dhcp server $DHCP scope $IPAddress add iprange $FirstIPAddress $LastIPAddress
&netsh dhcp server $DHCP scope $IPAddress set optionvalue 003 IPADDRESS $Gateway

 Switch -regex ($ScopeType) {
 {A} {
    $NetArea = "Admin"
    &netsh dhcp server $DHCP scope $IPAddress set optionvalue 252 STRING  "----"
    }
 {G} {
    $NetArea = "Guest"
    &netsh dhcp server $DHCP scope $IPAddress set optionvalue 006 IPADDRESS 8.8.8.8 8.8.4.4
    }
 {M} {
    $NetArea = "Mng"
    }
 {V} {
    $NetType = "Voice"
    &netsh dhcp server $DHCP scope $IPAddress set optionvalue 242 STRING  "----"
    }
 {W} {
    $NetType = "WiFi"
    &netsh dhcp server $DHCP scope $IPAddress set optionvalue 043 BYTE  "----"
    }
 }

 &netsh dhcp server $DHCP scope $IPAddress set state 1
 &netsh dhcp server $DHCP scope $IPAddress set comment $ScopeName
 &netsh dhcp server $DHCP scope $IPAddress set name $ScopeName

}