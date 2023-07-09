<#
.SYNOPSIS
    This function delete AppLocker settings using MDM WMI Bridge

.DESCRIPTION
    This script will delete AppLocker settings for EXE set using MDM WMI Bridge. Based on the script from Sandy Zeng: https://msendpointmgr.com/2020/09/20/does-applocker-work-in-windows-10-pro-yes-it-does/
    For more info see -> https://www.42dude.com/blog/2023/07/09/block-office-apps-using-powershell-and-applocker/
    
    Note: Use a tool to run the script as SYSTEM that applies at LOGOFF action of the end-user without an office license. I would suggest to run it also at logon for users with an office license to be sure all AppLocker settings are removed.

.VERSION HISTORY:
    1.0.0 - (2020-09-20) Script created by Sandy Zeng
    1.0.1 - (2023-07-09) AppLocker script - Delete AppLocker Settings
#>

$namespaceName = "root\cimv2\mdm\dmmap" #Do not change this
$className = "MDM_AppLocker_ApplicationLaunchRestrictions01_EXE03" #Do not change this
$GroupName = "AppLocker001" #Your own groupName
$parentID = "./Vendor/MSFT/AppLocker/ApplicationLaunchRestrictions/$GroupName"

Get-CimInstance -Namespace $namespaceName -ClassName $className -Filter "ParentID=`'$parentID`' and InstanceID='EXE'"  | Remove-CimInstance