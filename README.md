Delete a specific email from all Users Mailboxes in O365 and Exchange.
Powershell Script using Search-Mailbox to search and Delete Emails from user Mailboxes.

*Step 1 – Add your admin id to “Discovery Management” group
You can’t use Search-Mailbox cmd unless you have “Discovery Management” Group Rights, you can use EAC to Manage Role Group  and Grant permission to your account.

*Step 2 - To use -DeleteContent switch your account requires Mailbox Import Export role Access you can follow the below steps for the same. 

Create a group Name Import and Export
> New-RoleGroup “Import Export ” -Roles “Mailbox Import Export”
Add the admin ID to the role group.
> Add-RoleGroupMember “Import Export” -Member “LABADMIN” 
