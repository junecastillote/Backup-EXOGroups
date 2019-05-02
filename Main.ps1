#EDIT THESE VALUES
#1. Where is your script located?
$scriptDirectory = "C:\Scripts\Backup-ExoGroups"

#2. Where do we save the backup?
$backupDirectory = "C:\Scripts\Backup-ExoGroups\Backup"

#3. Where do we put the transcript log?
$logDirectory = "C:\Scripts\Backup-ExoGroups\Log"

#4. Which XML file contains your Office 365 Login?
#   If you don't have this yet, run this: Get-Credential | Export-CliXML <file.xml>
$credentialFile = "C:\Scripts\Backup-ExoGroups\credential.xml"

#5. If we will send the email summary, what is the sender email address we should use?
#   This must be a valid, existing mailbox and address in Office 365
#   The account you use for the Credential File must have "Send As" permission on this mailbox
$sender = "sender@domain.com"

#6. Who are the recipients?
#   Multiple recipients can be added (eg. "recipient1@domain.com","recipient2@domain.com")
$recipients = "june.castillote@gmail.com","june@lazyexchangeadmin.com"

#7. If you want to delete older backups, define the age in days.
$cleanBackupsOlderThanXDays = 60

#8. Should we compress the backup in a Zip file? $true or $false
$compressFiles = $true

#9. Do you want to send the email summary? $true or $false
$sendEmail = $true

#10. Do you want to backup the Distribution Groups and Members? $true or $false
$backupDistributionGroups = $true

#11. Do you want to backup the Dynamic DIstribution Groups? $true or $false
$backupDynamicDistributionGroups = $true
#------------------------------------------

#DO NOT TOUCH THE BELOW CODES
$params = @{
    backupDirectory = $backupDirectory
    logDirectory = $logDirectory
    credential = (Import-Clixml $credentialFile)
    sender = $sender
    recipients = $recipients
    cleanBackupsOlderThanXDays = $cleanBackupsOlderThanXDays
    compressFiles = $compressFiles
    sendEmail = $sendEmail
    backupDistributionGroups = $backupDistributionGroups
    backupDynamicDistributionGroups = $backupDynamicDistributionGroups
    #Limit = 10
}

."$scriptDirectory\Backup-EXOGroups.ps1" @params