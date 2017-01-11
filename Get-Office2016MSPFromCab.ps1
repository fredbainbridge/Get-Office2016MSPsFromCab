# Author - Fred Bainbridge
# Originally Published 5/31/2016
# 
# Usage Examples
#
# Use the defaut parameters in the script (update as needed)
# Get-Office2016MSPFromCab -UpdateDP $true -siteCode "LAB" -siteserver "cm01.cm.lab"
#
# Use explicit parameters, do not update the Distribution Points
# Get-Office2016MSPFromCab -SoftwareUpdatesFolder "\\cm01\SoftwareUpdates\Office2016x86Updates" -baseDestination "c:\fso1" -OfficeUpdatesFolder "\\cm01\Source\Microsoft Office 2016 x86\updates" 
#
# This assumes you have used ConfigMgr to download all the relevant Office 2016 patches and downloaded them to $SoftwareUpdatesFolder
 
#special thanks to https://technet.microsoft.com/en-us/magazine/2009.04.heyscriptingguy.aspx
Function ConvertFrom-Cab
{
 [CmdletBinding()]
 Param(
  $cab = "C:\fso\acab.cab",
  $destination
 )
 $comObject = "Shell.Application"
 Write-Verbose "Creating $comObject"
 $shell = New-Object -Comobject $comObject
 if(!$?) { $(Throw "unable to create $comObject object")}
 Write-Verbose "Creating source cab object for $cab"
 $sourceCab = $shell.Namespace("$cab").items()
 Write-Verbose "Creating destination folder object for $destination"
 if(-not (Test-Path $destination)) 
 {
    new-item $destination -ItemType Directory
 }
 $DestinationFolder = $shell.Namespace($destination)
 Write-Verbose "Expanding $cab to $destination"
 $DestinationFolder.CopyHere($sourceCab)
}

Function Get-Office2016MSPFromCab {
[CmdletBinding()]
param (
    $SoftwareUpdatesFolder = "\\cm01\Software Update Management\Microsoft Office 2016 x86 - Software Updates\",
    $baseDestination = "C:\fso1",
    $OfficeUpdatesFolder = "\\cm01\Source\Microsoft Office 2016 x86\updates",
    [bool]$UpdateDP = $false, 
    $siteCode = "LAB",
    $siteserver = "cm01.cm.lab",
    $appname = "Microsoft Office 2016 x86"

)

if(-not (test-path $baseDestination))
{
    new-item $baseDestination -ItemType Directory
}

#get all the cab files and copy them locally
write-host $SoftwareUpdatesFolder
Get-ChildItem -Path $SoftwareUpdatesFolder -Filter *.cab -Recurse | ForEach-Object {
    $GUID = (new-guid).Guid
    $destination = "$baseDestination\$GUID"
    Write-host  $destination
    ConvertFrom-Cab -cab $PSItem.FullName -destination $destination -Verbose
    
}

#rename and prepare the files for copy
Get-ChildItem -path $baseDestination -Filter *.msp -Recurse | ForEach-Object {
    Rename-Item -Path $PSItem.FullName -NewName ($PSItem.BaseName + (New-Guid).GUID + ".msp")
}
Get-ChildItem -Path $baseDestination -Filter *.msp -Recurse | Move-Item -Destination $baseDestination

#move the cabs to the Office Updates folder
Get-ChildItem -Path $baseDestination -Filter *.msp  | move-item -Destination $OffiCeUpdatesFolder -Verbose

#cleanup
Get-ChildItem -Path $baseDestination | Remove-Item -Recurse -Force -Verbose

#Update the content
if($UpdateDP){
(Get-Wmiobject -Namespace "root\SMS\Site_$sitecode" -Class SMS_ContentPackage -filter "Name='$appName'").Commit() 
}
}