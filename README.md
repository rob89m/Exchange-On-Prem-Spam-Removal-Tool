This is a simple PowerShell Form which leverage the Compliance Search function in Exchange to locate and remove SPAM emails that have been delivered to users mailboxes.

Using the tool is a simple
1. Launch the script from powershell as shown below
    .\SpamRemovalGUI.ps1 -ExchServer yourexchange.domain.com
2. Complete the relevant fields in the form that appears to search for emails in users mailboxes
3. Run the search
4. Review the emails found in the search (Most importantly at this stage confirm that the search did not return any emails that you do not wish to remove)
5. Pressing Remove will delete the emails shown on the list from the respective users mailboxes
 
Emails removed using this tool can still be seen and restored in Recover Deleted Items inside Outlook. If you wish to change this behaviour you can do so by changing the PurgeType on line 220.



This script is provided on an "As Is" basis with no guarantees. Anyone chosing to use this script does so at their own risk and should review the code in its entirity before running it in their own environment.
