### [Powershell Script] Using Search-Mailbox to remove Email from user Mailboxes based on Sender and Subject

No matter how much security you put in place, there are chances that your environment is hit by the Phishing and spam emails and you have received a request from your security team to clean that email from user mailboxes immediately.

### Example:Search and Report Only
> .\Search-Mailbox-and-Del-Email.ps1 -Subject "Insider -- April 13, 2017" -From "testuser@xyz.com" -InputFile .\user.txt

### Search and Delete Emails.
> .\Search-Mailbox-and-Del-Email.ps1 -Subject "Inside-- April 13, 2017" -From "testuser@test.com" -InputFile .\user.txt -DeleteEmails "True"

Read my post on this script for more Details:
https://www.linkedin.com/pulse/powershell-scriptusing-search-mailbox-remove-email-from-sunil-chauhan
