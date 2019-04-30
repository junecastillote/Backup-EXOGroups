
<#PSScriptInfo

.VERSION 1.0

.GUID cb824f8b-3a10-4cfe-a46d-a15e7bbf1358

.AUTHOR June Castillote

.COMPANYNAME www.lazyexchangeadmin.com

.COPYRIGHT june.castillote@gmail.com

.TAGS Office365 Script Backup DistributionGroup Dynamic Static Group PowerShell Tool Report Export

.LICENSEURI

.PROJECTURI

.ICONURI

.EXTERNALMODULEDEPENDENCIES 

.REQUIREDSCRIPTS

.EXTERNALSCRIPTDEPENDENCIES

.RELEASENOTES


.PRIVATEDATA

#>

<# 

.DESCRIPTION 
 Backup or Export Exchange Online Groups and Members to XML files.

#> 
Param(
        # Parameter help description
        [Parameter(Mandatory=$true)]
        [string]$loginlXML,
        [Parameter(Mandatory=$true)]
        [string]$configXML
)
#Functions------------------------------------------------------------------------------------------
#Function to connect to EXO Shell
Function New-EXOSession
{
    [CmdletBinding()]
    param(
        [parameter(mandatory=$true)]
        [PSCredential] $exoCredential
    )

    Get-PSSession | Remove-PSSession -Confirm:$false
    $EXOSession = New-PSSession -ConfigurationName "Microsoft.Exchange" -ConnectionUri 'https://ps.outlook.com/powershell' -Credential $exoCredential -Authentication Basic -AllowRedirection -WarningAction SilentlyContinue
    Import-PSSession $EXOSession -AllowClobber -DisableNameChecking | out-null
}

#Function to compress the CSV file
Function New-ZipFile
{
	[CmdletBinding()] 
    param ( 
        [Parameter(Mandatory)] 
        [string]$fileToZip,    
				[Parameter(Mandatory)]
				[string]$destinationZip
	)
	Add-Type -assembly System.IO.Compression
	Add-Type -assembly System.IO.Compression.FileSystem
	[System.IO.Compression.ZipArchive]$outZipFile = [System.IO.Compression.ZipFile]::Open($destinationZip, ([System.IO.Compression.ZipArchiveMode]::Create))
	[System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile($outZipFile, $fileToZip, (Split-Path $fileToZip -Leaf)) | out-null
	$outZipFile.Dispose()
}

#Function to delete old files based on age
Function Invoke-Housekeeping
{
    [CmdletBinding()] 
    param ( 
        [Parameter(Mandatory)] 
        [string]$folderPath,
    
		[Parameter(Mandatory)]
		[int]$daysToKeep
    )
    
    $datetoDelete = (Get-Date).AddDays(-$daysToKeep)
    $filesToDelete = Get-ChildItem $FolderPath | Where-Object { $_.LastWriteTime -lt $datetoDelete }

    if (($filesToDelete.Count) -gt 0) {	
		foreach ($file in $filesToDelete) {
            Remove-Item -Path ($file.FullName) -Force -ErrorAction SilentlyContinue
		}
	}	
}

#Function to Stop Transaction Logging
Function Stop-TxnLogging
{
	$txnLog=""
	Do {
		try {
			Stop-Transcript | Out-Null
		} 
		catch [System.InvalidOperationException]{
			$txnLog="stopped"
		}
    } While ($txnLog -ne "stopped")
}

#Function to Start Transaction Logging
Function Start-TxnLogging 
{
    param 
    (
        [Parameter(Mandatory=$true)]
        [string]$logPath
    )
	Stop-TxnLogging
    Start-Transcript $logPath -Append
}
#----------------------------------------------------------------------------------------------------
Stop-TxnLogging
Clear-Host
$scriptVersion = (Test-ScriptFileInfo -Path $MyInvocation.MyCommand.Definition).version

$script_root = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition
[xml]$config = Get-Content "$($script_root)\config.xml"
#debug-----------------------------------------------------------------------------------------------
$enableDebug = $config.options.enableDebug
#----------------------------------------------------------------------------------------------------
#backup group types----------------------------------------------------------------------------------
$backupDistributionGroups = $config.options.backupDistributionGroups
$backupDynamicDistributionGroups = $config.options.backupDynamicDistributionGroups
#----------------------------------------------------------------------------------------------------
#Mail------------------------------------------------------------------------------------------------
$sendReport = $config.options.SendReport
[string]$tenantName = $config.options.tenantName
[string]$fromAddress = $config.options.fromAddress
[string]$toAddress = $config.options.toAddress
[string]$smtpServer = "smtp.office365.com"
[int]$smtpPort = "587"
[string]$mailSubject = "Exchange Online Groups Backup"
#----------------------------------------------------------------------------------------------------
#Housekeeping----------------------------------------------------------------------------------------
$enableHousekeeping = $config.options.enableHousekeeping

if (!$config.options.daysToKeep)
{
    [int]$daysToKeep = 1
}
else 
{
    [int]$daysToKeep = $config.options.daysToKeep
}
#----------------------------------------------------------------------------------------------------
$Today=Get-Date
[string]$fileSuffix = '{0:dd-MMM-yyyy_hh-mm_tt}' -f $Today
$logPath = "$($script_root)\Logs"
$logFile = "$($logPath)\DebugLog_$($fileSuffix).txt"
$backupDir = "$($script_root)\BackupDir"
$backupPath = "$($script_root)\BackupDir\$($fileSuffix)"
$DG_backupFile = "$($backupPath)\DistributionGroups.xml"
$DDG_backupFile = "$($backupPath)\DynamicDistributionGroups.xml"
$zipFile = "$($backupDir)\Backup_$($fileSuffix).zip"

#Create folders if not found
if (!(Test-Path $logPath)) {New-Item -ItemType Directory -Path $logPath | Out-Null}
if (!(Test-Path $backupPath)) {New-Item -ItemType Directory -Path $backupPath | Out-Null}



#start transcribing----------------------------------------------------------------------------------
if ($enableDebug) {Start-Transcript -Path $logFile}
#----------------------------------------------------------------------------------------------------

#check if credential.xml is present------------------------------------------------------------------
#if (!(Test-Path "$($script_root)\credential.xml"))
#{
#    Write-Host (get-date -Format "dd-MMM-yyyy hh:mm:ss tt") ": The file 'credential.xml' does not exist. Please run 'Get-Credential | Export-CliXml .\credential.xml' first to setup the authorization account for Office 365" -ForegroundColor RED
#    EXIT
#}
#----------------------------------------------------------------------------------------------------

#BEGIN------------------------------------------------------------------------------------------
Write-Host (get-date -Format "dd-MMM-yyyy hh:mm:ss tt") ": Begin" -ForegroundColor Green
Write-Host (get-date -Format "dd-MMM-yyyy hh:mm:ss tt") ": Connecting to Exchange Online Shell" -ForegroundColor Green

#Connect to O365 Shell
#Note: This uses an encrypted credential (XML). To store the credential:
#1. Login to the Server/Computer using the account that will be used to run the script/task
#2. Run this "Get-Credential | Export-CliXml credential.xml"
#3. Make sure that credential.xml is in the same folder as the script.
$credential = Import-Clixml "$($script_root)\credential.xml"
try 
{
    New-EXOSession $credential
}
catch 
{
    Write-Host (get-date -Format "dd-MMM-yyyy hh:mm:ss tt") ": There was an error connecting to Exchange Online. Terminating Script" -ForegroundColor YELLOW
    Stop-TxnLogging
    EXIT
}

#Start Export Process---------------------------------------------------------------------------
if ($backupDistributionGroups)
{
    $members = @()
    Write-Host (get-date -Format "dd-MMM-yyyy hh:mm:ss tt") ": Retrieving Distribution Groups" -ForegroundColor Yellow
    $dgrouplist = Get-DistributionGroup -ResultSize Unlimited -WarningAction SilentlyContinue | Sort-Object Name
    Write-Host (get-date -Format "dd-MMM-yyyy hh:mm:ss tt") ": There are a total of $($dgrouplist.count) Distribution Groups" -ForegroundColor Yellow
    $dgrouplist | Export-Clixml -Path $DG_backupFile -Depth 5

    $i=1
    foreach ($dgroup in $dgrouplist)
    {
        #[array]$dgroup_members = Get-DistributionGroupMember -id $dgroup.DistinguishedName -ResultSize Unlimited
        #Write-Host (get-date -Format "dd-MMM-yyyy hh:mm:ss tt") ": ($($i) of $($dgrouplist.count)) | [$($dgroup_members.count) members] | $($dgroup.DisplayName)" -ForegroundColor Yellow
        #$dgroup_members | Export-Clixml -Depth 5 -Path "$backupPath\$($dgroup.PrimarySMTPAddress)_dgMembers.xml"
        $dgroup_members = Get-Group $dgroup.DistinguishedName | Select-Object Members,WindowsEmailAddress,DisplayName
        Write-Host (get-date -Format "dd-MMM-yyyy hh:mm:ss tt") ": ($($i) of $($dgrouplist.count)) | [$($dgroup_members.members.count) members] | $($dgroup.DisplayName)" -ForegroundColor Yellow
        #$dgroup_members | Export-Clixml -Depth 5 -Path "$backupPath\$($dgroup.WindowsEmailAddress).xml"
        $i=$i+1
        $members += $dgroup_members
    }
    $members | Export-Clixml -Path "$backupPath\DistributionGroupMembers.xml"
}

if ($backupDynamicDistributionGroups)
{
    Write-Host (get-date -Format "dd-MMM-yyyy hh:mm:ss tt") ": Retrieving Dynamic Distribution Groups" -ForegroundColor Yellow
    $ddgrouplist = Get-DynamicDistributionGroup -ResultSize Unlimited -WarningAction SilentlyContinue | Sort-Object Name
    Write-Host (get-date -Format "dd-MMM-yyyy hh:mm:ss tt") ": There are a total of $($ddgrouplist.count) Dynamic Distribution Groups" -ForegroundColor Yellow
    $ddgrouplist | Export-Clixml -Path $DDG_backupFile -Depth 5
}
#----------------------------------------------------------------------------------------------------

#Zip the file to save space---------------------------------------------------------------------
Compress-Archive -Path "$backupPath\*.*" -DestinationPath $zipFile -CompressionLevel Optimal
Write-Host (get-date -Format "dd-MMM-yyyy hh:mm:ss tt") ": Backup Saved to $zipFile" -ForegroundColor Yellow
$zipSize = (Get-ChildItem $zipFile | Measure-Object -Property Length -Sum)
#Allow some time (in seconds) for the file access to close, increase especially if the resulting files are huge, or server I/O is busy.
$sleepTime=5
Write-Host (get-date -Format "dd-MMM-yyyy hh:mm:ss tt") ": Pending next operation for $($sleepTime) seconds" -ForegroundColor Yellow
Start-Sleep -Seconds $sleepTime
Remove-Item -Path $backupPath -Recurse -Force
#----------------------------------------------------------------------------------------------------

#Invoke Housekeeping----------------------------------------------------------------------------
if ($enableHousekeeping)
{
	Write-Host (get-date -Format "dd-MMM-yyyy hh:mm:ss tt") ": Deleting backup files older than $($daysToKeep) days" -ForegroundColor Yellow
	Invoke-Housekeeping -folderPath $backupDir -daysToKeep $daysToKeep
}
#-----------------------------------------------------------------------------------------------
#Count the number of backups existing and the total size----------------------------------------
$backupFiles = (Get-ChildItem $backupDir | Measure-Object -Property Length -Sum)
#-----------------------------------------------------------------------------------------------
$timeTaken = New-TimeSpan -Start $Today -End (Get-Date)
#Send email if option is enabled ---------------------------------------------------------------
if ($SendReport)
{
Write-Host (get-date -Format "dd-MMM-yyyy hh:mm:ss tt") ": Sending report to" ($toAddress -join ";") -ForegroundColor Yellow
$xSubject="[$($tenantName)] $($mailSubject): " + ('{0:dd-MMM-yyyy hh:mm:ss tt}' -f $Today)
$htmlBody=@'
<!DOCTYPE html>
<html>
<head>
<style>
table {
  font-family: "Century Gothic", sans-serif;
  border-collapse: collapse;
  width: 100%;
}
td, th {
  border: 1px solid #dddddd;
  text-align: left;
  padding: 8px;
}

</style>
</head>
<body>
<table>
'@
$htmlBody+="<tr><th>----SUMMARY----</th></tr>"
$htmlBody+="<tr><th>Number of Distribution Groups</th><td>$($dgrouplist.count)</td></tr>"
$htmlBody+="<tr><th>Number of Dynamic Distribution Groups</th><td>$($ddgrouplist.count)</td></tr>"
$htmlBody+="<tr><th>Backup Server</th><td>"+(Get-Content env:computername)+"</td></tr>"
$htmlBody+="<tr><th>Backup File</th><td>$($zipFile)</td></tr>"
$htmlBody+="<tr><th>Backup Size</th><td>"+ ("{0:N0}" -f ($zipSize.Sum / 1KB)) + " KB</td></tr>"
$htmlBody+="<tr><th>Time to Complete</th><td>"+ ("{0:N0}" -f $($timeTaken.TotalMinutes)) + " Minutes</td></tr>"
$htmlBody+="<tr><th>Total Number of Backups</th><td>$($backupFiles.Count)</td></tr>"
$htmlBody+="<tr><th>Total Backup Folder Size</th><td>"+ ("{0:N2}" -f ($backupFiles.Sum / 1KB)) + " KB</td></tr>"
$htmlBody+="<tr><th>----SETTINGS----</th></tr>"
$htmlBody+="<tr><th>Tenant Organization</th><td>$($tenantName)</td></tr>"
$htmlBody+="<tr><th>Debug Enabled</th><td>$($enableDebug)</td></tr>"
$htmlBody+="<tr><th>Housekeeping Enabled</th><td>$($enableHousekeeping)</td></tr>"
$htmlBody+="<tr><th>Days to Keep</th><td>$($daysToKeep)</td></tr>"
$htmlBody+="<tr><th>Report Recipients</th><td>" + $toAddress.Replace(",","<br>") + "</td></tr>"
$htmlBody+="<tr><th>SMTP Server</th><td>$($smtpServer)</td></tr>"
$htmlBody+="<tr><th>Script Path</th><td>$($MyInvocation.MyCommand.Definition)</td></tr>"
$htmlBody+="<tr><th>Script Source Site</th><td><a href=""https://github.com/junecastillote/Export-O365GroupsAndMembers"">Export-O365GroupsAndMembers.ps1</a> version $($scriptVersion)</td></tr>"
$htmlBody+="</table></body></html>"
Send-MailMessage -from $fromAddress -to $toAddress.Split(",") -subject $xSubject -body $htmlBody -dno onSuccess, onFailure -smtpServer $SMTPServer -Port $smtpPort -Credential $credential -UseSsl -BodyAsHtml
}
#-----------------------------------------------------------------------------------------------
Write-Host (get-date -Format "dd-MMM-yyyy hh:mm:ss tt") ": End" -ForegroundColor Green
#-----------------------------------------------------------------------------------------------
Stop-TxnLogging
