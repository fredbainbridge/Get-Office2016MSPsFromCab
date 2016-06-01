# Author - Fred Bainbridge
# Originally Published 5/31/2016
# 

$cabsFolder = "\\cm01\Software Update Management\Office 2016 x64 Updates\"
$baseDestination = "C:\fso1"
$OffiCeUpdatesFolder = "\\cm01\Source\Microsoft Office 2016 x86\updates"

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


#get all the cab files and copy them locally
Get-ChildItem -Path $cabsFolder -Filter *.cab -Recurse | ForEach-Object {
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

