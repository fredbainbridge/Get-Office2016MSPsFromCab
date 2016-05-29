
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


#$cab = "\\cm\Software Update Management\Office 2016 x64 Updates\0b29b011-b8fe-4051-ab53-016ac266e1db\outlook-x-none.cab"
$cabsFolder = "\\cm\Software Update Management\Office 2016 x64 Updates\"
$baseDestination = "C:\fso1"

#get all the cabs
Get-ChildItem -Path $cabsFolder -Filter *.cab -Recurse | ForEach-Object {
    $GUID = (new-guid).Guid
    $destination = "$baseDestination\$GUID"
    Write-host  $destination
    ConvertFrom-Cab -cab $PSItem.FullName -destination $destination -Verbose
    
}


Get-ChildItem -path $baseDestination -Filter *.msp -Recurse | ForEach-Object {
    Rename-Item -Path $PSItem.FullName -NewName ($PSItem.BaseName + (New-Guid).GUID + ".msp")
}

Get-ChildItem -Path $baseDestination -Filter *.msp -Recurse | Move-Item -Destination $baseDestination

#move the cabs to the Office Updates folder
Get-ChildItem -Path C:\fso1 -Filter *.msp  | move-item -Destination $OffiCeUpdatesFolder

#cleanup
Get-ChildItem -Path c:\fso1 | Remove-Item -Recurse -Force

