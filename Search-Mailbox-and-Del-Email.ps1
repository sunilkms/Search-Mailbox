# Example:Search and Report Only
# .\Search-Mailbox-and-Del-Email.ps1 -Subject "Inside Shire -- April 13, 2017" -From "testuser@xyz.com" -InputFile .\user.txt
#
# Search and Delete Emails.
# .\Search-Mailbox-and-Del-Email.ps1 -Subject "Inside Shire -- April 13, 2017" -From "testuser@test.com" -InputFile .\user.txt -DeleteEmails "True"
#

Param(
$From="spammer@spammer.com",
$Subject="Spam",
$TargetMailbox=$Cred.UserName,
$DeleteEmails,
$InputFile
)

$fromst="FROM:`"$from`""
$subjectst="Subject:`"$Subject`""
$SQ="$fromst AND $subjectst"
$ExportReport="True"
$reportOnly="True"
$SearchResultFolderName="Search-Mailbox-Results-$(Get-date -f dd-MM-yy)"
$ReportFileName="Search-Mailbox-Results-$(Get-date -f dd-MM-yy).csv"

#Save Mailbox Details where Emails Was Found
$report=@()
if($deleteEmails){
        $reportOnly="False" 
        if ($InputFile)
        {
         Write-Host "Emails Found in Search Criteria will be deleted" -ForegroundColor Cyan
         ""
         Write-Host "Validating Mailbox for users provided in the input file"
         ""
         $ValidatedMailbox=@()
         foreach ($Mbx in $(Get-Content $InputFile))
         {
         try{
         Write-Host "Checking Mailbox:"$mbx
         $mbx = Get-Mailbox $mbx -ea SilentlyContinue
         if($mbx){$ValidatedMailbox+=$mbx}
          }
         catch{ Write-Host "Mailbox not found:"$mbx -f Yellow }         
         }
         ""
         Write-host "Email Will be searched and Deleted against the below users." -ForegroundColor Yellow
         $ValidatedMailbox
         ""                  
         $SR=$ValidatedMailbox | Search-Mailbox -SearchQuery $SQ -TargetMailbox $TargetMailbox `
         -TargetFolder $SearchResultFolderName -LogLevel Full -DeleteContent
         if ($SR.ResultItemsCount -gt 0) {$MbxhasMail+=$EmailAdrs ; $report+=$SR}                
         }#Delete emails for All users

Else {write-host "No input file provided, for searching in All Mailbox use -SearchAllMbx True paramiter" -ForegroundColor Cyan}
   }

if($reportOnly -eq "True") {
#Delete emails for specific users
if ($InputFile)
        {
         ""
         Write-Host "Validated Mailbox for users provided in the input file."
         ""
         $ValidatedMailbox=@()
         foreach ($Mbx in $(Get-Content $InputFile))
         {
         try{
         Write-Host "Checking Mailbox:"$mbx
         $mbx = Get-Mailbox $mbx -ea SilentlyContinue
         if($mbx){$ValidatedMailbox+=$mbx}
          }
         catch{ Write-Host "Mailbox not found:"$mbx -f Yellow }         
         }
         $ValidatedMailbox
         ""
         Write-host "Searchign emails for the provided users..." -ForegroundColor Yellow        
         $SR=$ValidatedMailbox | Search-Mailbox -SearchQuery $SQ -TargetMailbox $TargetMailbox -TargetFolder $SearchResultFolderName -LogLevel Full
         if ($SR.ResultItemsCount -gt 0) {$MbxhasMail+=$EmailAdrs ; $report+=$SR}                
         } #Delete emails for All users
Else { write-host "No input file provided, for searching in All Mailbox use -SearchAllMbx True paramiter" -ForegroundColor Cyan }
}

#Block For Deleting All Emails             
if ($ExportReport) {    
    if ($report) {
    $report = $report | select Identity,TargetMailbox,TargetFolder,ResultItemsCount,ResultItemsSize
    #$report | ft
    $report | Export-csv $ReportFileName -NoTypeInformation 
    Write-host "Report Exported and saved at" $($(Get-Location).Path + "\" + $ReportFileName)
    } else { "No Data found for export" }
}
