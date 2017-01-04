[CmdletBinding()]
param(
    [string]$siteserver = "localhost",
    [string]$sitecode = "LAB",
    [string]$StagingLocation = "c:\fso1",
    [string]$OfficeInstallatioNSourcePath = "\\cm01\Source\Microsoft Office 2016 x86\",
    [string]$OfficeUpdatesFile = "https://raw.githubusercontent.com/fredbainbridge/Get-Office2016MSPsFromCab/master/Office2016-Oct2016-SoftwareUpdates.txt",
    [bool]$UpdateDP = $false, 
    [string]$appname = "Microsoft Office 2016 x86"
)
$NameSpace = "root\SMS\Site_$sitecode"
$class = "SMS_SoftwareUpdate"
    
#download the text file.
$updates = (Invoke-WebRequest -Uri $OfficeUpdatesFile).content

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
$UpdateLine = @();

#debug
#$FileName = "Office2016OctUpdates-Debug.txt"

$Updates -split '[\r\n]' |? {$_}| ForEach-Object {
    $UpdateName, $URL, $FileName = $PSItem.split(",")
    
    $FileName = $StagingLocation + "\" +  (New-Guid).GUID + $FileName
       
    Start-BitsTransfer -Source $URL -Destination $FileName
    If(Test-Path $FileName)
    {
        $GUID = (new-guid).Guid
        $destination = "$StagingLocation\$GUID"
        ConvertFrom-Cab -cab $FileName -destination $destination
        Remove-Item -Path $FileName
        Get-ChildItem -path $destination -Filter *.msp -Recurse | ForEach-Object {
            Rename-Item -Path $PSItem.FullName -NewName ($PSItem.BaseName + (New-Guid).GUID + ".msp")
        }
        Get-ChildItem -Path $destination -Filter *.msp -Recurse | Move-Item -Destination $StagingLocation
        Remove-Item -path $destination -Recurse -Force
    }
}


#delete existing updates - be careful here.  move any custom msps from this location first.
Get-ChildItem -Path "$OfficeInstallatioNSourcePath\updates" -Filter *.msp | Remove-Item -Force 

#move update to office folder
Get-ChildItem -Path $StagingLocation -Filter *.msp  | move-item -Destination "$OfficeInstallatioNSourcePath\updates" -Verbose

#Update the content
#if($UpdateDP){
#    (Get-Wmiobject -Namespace "root\SMS\Site_$sitecode" -Class SMS_ContentPackage -filter "Name='$appName'").Commit() 
#}
import-module 'C:\Program Files (x86)\Microsoft Configuration Manager\AdminConsole\bin\ConfigurationManager.psd1' -force #make this work for you
if ((get-psdrive $sitecode -erroraction SilentlyContinue | measure).Count -ne 1) {
    new-psdrive -Name $SiteCode -PSProvider "AdminUI.PS.Provider\CMSite" -Root $SiteServer
}
set-location $sitecode`:

$dptypes = Get-CMDeploymentType -ApplicationName "$appname"
foreach ($dpt in $dptypes)
{
    $dptname = $dpt.LocalizedDisplayName
    Write-Verbose "Deployment Type: $dptname"
    Update-CMDistributionPoint -ApplicationName "$appname" -DeploymentTypeName "$dptname"
}

           
#see example here 
#https://social.technet.microsoft.com/Forums/systemcenter/en-US/f11a43e0-409c-443a-adb0-74de102c40f7/add-updates-to-a-deployment-package-using-powershell?forum=configmgrgeneral&prof=required


#Get-WmiObject -Class $class -Namespace $name