See https://www.42dude.com/blog/2023/07/09/block-office-apps-using-powershell-and-applocker/ for more info

1: Apply the powershell script "Enable-Applocker.ps1" at LOGON for the targetted users and run as SYSTEM
2: Apply the powershell script "Delete-Applocker-Policies.ps1" at LOGOFF for the targetted users and run as SYSTEM (tip: also run at LOGON for the users not targetted at step 1)
