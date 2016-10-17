$siteserver = "localhost"
$sitecode = "LAB"
$NameSpace = "root\SMS\Site_$sitecode"
$StagingLocation = "c:\fso1"
$OfficeInstallatioNSourcePath = "\\cm01\Software Update Management\Office2016x86"
$class = "SMS_SoftwareUpdate"
$FileName = "Office2016-Oct2016-SoftwareUpdates.txt"
<#
#create new software updates package
$WMIConnection = [WMICLASS]"\\$SiteServer\ROOT\SMS\Site_$($sitecode):SMS_SoftwareUpdatesPackage"
$NewSUPPackage = $WMIConnection.psbase.CreateInstance()
$NewSUPPackage.Name = "PowerShell SUP Package"
$NewSUPPackage.Description = "PowerShell TEST"
$NewSUPPackage.PkgSourceFlag = 2
$NewSUPPackage.PkgSourcePath = "$path"
$NewSUPPackage.Put()
 
$SoftwareUpdatePackage = Get-WmiObject -Class SMS_SoftwareupdatesPackage -Namespace $name
#>
#download the files.

#debug
#$FileName = "Office2016OctUpdates-Debug.txt"
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

Get-Content $FileName | ForEach-Object {
    $CI_ID = (Get-WmiObject -Class $class -Namespace $NameSpace -Filter "LocalizedDisplayName='$PSItem'" -Property "CI_ID").CI_ID
    $ContentID = (get-wmiobject -Query "select * from SMS_CItoContent where ci_id=$CI_ID" -Namespace $NameSpace).ContentID
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
                    Rename-Item -Path $PSItem.FullName -NewName ($PSItem.BaseName + (New-Guid).GUID + ".msp")
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



#see example here 
#https://social.technet.microsoft.com/Forums/systemcenter/en-US/f11a43e0-409c-443a-adb0-74de102c40f7/add-updates-to-a-deployment-package-using-powershell?forum=configmgrgeneral&prof=required


#Get-WmiObject -Class $class -Namespace $name