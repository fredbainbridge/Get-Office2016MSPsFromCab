# The Update must be synced in Configmgr for this to work.  They don't have to be downloaded.
# $OfficeUpdatesFile is a list of software updates LocalizedDisplayName to download.
# The downloads are cab files, the msp files are then extracted from the cab files and saved to $StagingLocation
# The $OfficeUpdatesFile is located in my github repository and will be be updated periodically. You can maintain this list yourself, just update this parameter. 
# if the file is local, comment out line 40, 41 and 44 and uncomment out line 45 

[CmdletBinding()]
param(
    [string]$siteserver = "localhost",
    [string]$sitecode = "LAB",
    [string]$StagingLocation = "c:\fso1",
    [string]$OfficeUpdatesFile = "https://raw.githubusercontent.com/fredbainbridge/Get-Office2016MSPsFromCab/master/Office2016-SoftwareUpdates.txt"
)
$class = "SMS_SoftwareUpdate"
$NameSpace = "root\SMS\Site_$sitecode"

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
#one off downloads
#$item = "Security Update for Microsoft Word 2016 (KB3128057) 32-Bit Edition"
#$item | ForEach-Object {

$updates = (Invoke-WebRequest -Uri $OfficeUpdatesFile).content
$Updates -split '[\r\n]' |? {$_} |  ForEach-Object { 
#Get-Content $OfficeUpdatesFile | ForEach-Object {
    $KB = ($PSITEM -replace "^.*?(?=KB)", "") -replace "\W(.*)", ""
    Write-Host "$KB - Downloading Update - $PSITEM"
    
    $CI_ID = (Get-WmiObject -ComputerName $siteserver -Class $class -Namespace $NameSpace -Filter "LocalizedDisplayName='$PSItem'" -Property "CI_ID").CI_ID
    $ContentID = (get-wmiobject -ComputerName $siteserver -Query "select * from SMS_CItoContent where ci_id=$CI_ID" -Namespace $NameSpace).ContentID
    #get the content location (URL)
    $ContentID | ForEach-Object {
        $objContent = Get-WmiObject -ComputerName $siteserver -Namespace $NameSpace -Class SMS_CIContentFiles -Filter "ContentID = '$PSITEM'" 
        $FileName = $StagingLocation + "\" +  (New-Guid).GUID + $objContent.FileName
        $URL = $objContent.SourceURL
        try 
        {
            Start-BitsTransfer -Source $URL -Destination $FileName
            If(Test-Path $FileName)
            {
                $GUID = (new-guid).Guid
                $destination = "$StagingLocation\$GUID"
                ConvertFrom-Cab -cab $FileName -destination $destination
                Remove-Item -Path $FileName
                Get-ChildItem -path $destination -Filter *.msp -Recurse | ForEach-Object {
                    Rename-Item -Path $PSItem.FullName -NewName ("$kb-" + (New-Guid).GUID + ".msp")
                }
                Get-ChildItem -Path $destination -Filter *.msp -Recurse | Move-Item -Destination $StagingLocation
                Remove-Item -path $destination -Recurse -Force
            }
        }
        catch
        {
            write-host "stopping here"
        }
    }
}