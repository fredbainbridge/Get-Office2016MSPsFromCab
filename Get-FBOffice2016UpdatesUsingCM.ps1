#this downloads the information about updates given the localized display name and stores it in a variable.  
#This can be used to create the software update detailed text file. (name, url and filename)
[CmdletBinding()]
param(
    $siteserver = "localhost",
    $sitecode = "LAB",
    $StagingLocation = "c:\fso1",
    $OfficeInstallatioNSourcePath = "\\cm01\Software Update Management\Office2016x86",
    $OfficeUpdatesFile = "https://raw.githubusercontent.com/fredbainbridge/Get-Office2016MSPsFromCab/master/Office2016OctUpdates-LocalizedName.txt"
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


$Updates -split '[\r\n]' |? {$_}| ForEach-Object {
    $UpdateName = $PSItem
    $CI_ID = (Get-WmiObject -Class $class -Namespace $NameSpace -Filter "LocalizedDisplayName='$UpdateName'" -Property "CI_ID").CI_ID
    $ContentID = (get-wmiobject -Query "select * from SMS_CItoContent where ci_id=$CI_ID" -Namespace $NameSpace).ContentID
    #get the content location (URL)
    $ContentID | ForEach-Object {
        $objContent = Get-WmiObject -ComputerName $siteserver -Namespace $NameSpace -Class SMS_CIContentFiles -Filter "ContentID = '$PSITEM'"  
        $FileName = $objContent.FileName
        $URL = $objContent.SourceURL
        $UpdateLine += "$UpdateName,$URL,$FileName"
    }
}

#$UpdateLine | clip

$UpdateLine > Office2016-Oct2016-SoftwareUpdates.txt

#The git stuff
git add .\Office2016-Oct2016-SoftwareUpdates.txt
git add .\Office2016OctUpdates-LocalizedName.txt
git commit -a -m "updated Office KB files"
git push

.\Get-FBOffice2016Updates.ps1

#$UpdateLine | ForEach-Object {$psItem.split(",")[0]} | select -Unique | clip
