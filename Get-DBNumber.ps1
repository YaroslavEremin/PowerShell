Function Get-DBNumber {
<#
  .SYNOPSIS
  This function returns the first pass in the numbering of mail databases Exchange
  .DESCRIPTION
  The search is performed by isolating the mask numbers of databases. Then construct a sorted list and Ichetu first pass in the numbering
  .EXAMPLE
   $n = Get-DBNumber
  #>
$NumbersArray = @()
[int]$Counter = 0
$DBArray = (Get-MailboxDatabase | Where-Object -Property Name -like -Value "DB*" | Select-Object Name)
ForEach ($DBName in $DBArray) {
$Number = [regex]::match($DBName.Name,'\d{2,3}').Value
$NumbersArray += ($Number -as [int])
}
$NumbersArray = $NumbersArray | Sort-Object
While (( $NumbersArray[$Counter] - $Counter) -le 1) {
$Counter++
}
$DBNumber = ($Counter + 1) -as [string]
if ($DBNumber.Length -eq 1) {
$DBNumber = "0$DBNumber"
}
Return "$DBNumber"
}