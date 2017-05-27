# Search Emails based on Sender and Subject to delete emails from user mailboxes.
# Author: Sunil Chauhan 
# My Blog : http://www.sunilchauhan.info
# Email Me @ Sunilkms@gmail.com 

Param(
$fromString="jcordeiro@sunilchauhan.info",
$SubjectString="HRConnect is Here for You",
$adminuser=$Cred.UserName,
$deleteEmails,
$MbxFilePath,
$SearchAllMbx,
$SearchAllMbxAndDelete,
$SearchResultFolderName="Search-Mailbox-Results-$(Get-date -f dd-MM-yy)",
$reportFineName="Search-Mailbox-Results-$(Get-date -f dd-MM-yy).csv"
)

$from="FROM:`"$fromString`""
$subject="Subject:`"$SubjectString`""
$SQ="$from AND $subject"

#Save Mailbox Details where Emails Was Found
$Global:MbxhasMail=@()
$report=@()

#Delete emails for specific users
if ($MbxFilePath)
        {
         Write-host "Checking if emails exist in specific users"   
            foreach ($EmailAdrs in $mbxfilepath) 
                {
                $SR=Search-Mailbox -Identity $EmailAdrs -SearchQuery $SQ `
                -TargetMailbox $Adminuser -TargetFolder $SearchResultFolderName -LogLevel Full
                if ($SR.ResultItemsCount -gt 0) {$MbxhasMail+=$EmailAdrs ; $report+=$SR}
                }
        }#Delete emails for All users
Else {write-host "No input file provided, for searching in All Mailbox use -SearchAllMbx True paramiter" -ForegroundColor Cyan}

if($deleteEmails){
                   if($MbxhasMail)
                   {
                   Write-Warning "Note:Emails matching search Criteria will be removed from user Mailbox, `
                   if you wish to just report if Emails are present in user mailbox run Script without -DeleteEmails param."
                   Foreach ($user in $MbxhasMail) 
                       {
                       Write-Host "Email Will be deleted from the user mailbox" -ForegroundColor Cyan
                       $SR=Search-Mailbox -Identity $EmailAdrs -SearchQuery $SQ `
                       -TargetMailbox $Adminuser -TargetFolder $SearchResultFolderName -LogLevel Full -DeleteContent
                       }
                   }
                   else{ Write-Host "No user found to delete emails" -f Yellow }                
                 }
                 
#Block For Deleting All Emails
                          
if ($SearchAllMbxAndDelete) {
        Read-host "Email Will Be deleted from user mailbox if found, please hit enter to continue..."
        #$mailbox = Get-Mailbox -resultSize Unlimited 
        $mailbox = Get-Mailbox -resultSize 500
        foreach ($mbx in $mailbox) 
                {
                $SR=Search-Mailbox -Identity $mbx.alias -SearchQuery $SQ `
                -TargetMailbox $Adminuser -TargetFolder $SearchResultFolderName -LogLevel Full
                $report+=$SR
                }
        }

if ($SearchAllMbx) {
        Read-host "This will just Report the if email found in user mailbox, please hit enter to continue..."
        #$mailbox = Get-Mailbox -resultSize Unlimited 
        $mailbox = Get-Mailbox -resultSize 500
        foreach ($mbx in $mailbox) 
                {
                $SR=Search-Mailbox -Identity $mbx.alias -SearchQuery $SQ `
                -TargetMailbox $Adminuser -TargetFolder $SearchResultFolderName -LogLevel Full
                $report+=$SR
                }
        }

$report = $report | select Identity,TargetMailbox,TargetFolder,ResultItemsCount,ResultItemsSize
$report | Export-csv $ReportFileName -NoTypeInformation
