# AD-O365-TerminationAutomation
Script used for automating the termination process of users from on-prem active directory, and Office 365.

The script assumes the following:
-The on-prem and O365 environments are dir synced.
-The employeeOU, shortTermOU, and longTermOU OUs are synced to O365, while the termedOU is not.
-The serviceAccount has the appropriate authorization to perform all tasks in the script.
-extensionAttribute2 is being used for the terimation date, and it's value is formatted in a common date format.
-The cred.txt file is storing the encrypted password of the serviceAccount, and was created with steps similar to those outlined here: https://blogs.technet.microsoft.com/robcost/2008/05/01/powershell-tip-storing-and-using-password-credentials/
-The comment attribute is storing a legal hold flag with a value of "lithold".

I'm still adding to this, including greater logging and email notifications. However, others have mentioned wanting to see the code so far.
