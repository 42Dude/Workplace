<#
.SYNOPSIS
    AppLocker configuration for restricting Microsoft Office (or other executables)

.DESCRIPTION
    This script will create AppLocker settings for EXE and restrict Microsoft Office. Based on the script from Sandy Zeng: https://msendpointmgr.com/2020/09/20/does-applocker-work-in-windows-10-pro-yes-it-does/.
    For more info see -> https://www.42dude.com/blog/2023/07/09/block-office-apps-using-powershell-and-applocker/

    NOTE: Use a tool to run the script as SYSTEM that applies at LOGON action of the end-user without an office license.

    Dont forget to delete the configuration when logging out :).

.VERSION HISTORY
    1.0.0 - (2020-09-20) Script created by Sandy Zeng
    1.0.1 - (2023-07-09) AppLocker script - No Office for Regular Users
#>

$namespaceName = "rootcimv2mdmdmmap" #Do not change this
$className = "MDM_AppLocker_ApplicationLaunchRestrictions01_EXE03" #Do not change this
$GroupName = "AppLocker001" #You can use your own Groupname, don't use special charaters or with space
$parentID = "./Vendor/MSFT/AppLocker/ApplicationLaunchRestrictions/$GroupName"

Add-Type -AssemblyName System.Web

#This is example Rule Collection for EXE, you should change this to your own settings
$obj = [System.Net.WebUtility]::HtmlEncode(@"
<RuleCollection Type="Exe" EnforcementMode="NotConfigured">
<FilePathRule Id="fd686d83-a829-4351-8ff4-27c7de5755d2" Name="(ADMIN) All files allowed" Description="Users of the Local Administrator group still have unrestricted access." UserOrGroupSid="S-1-5-32-544" Action="Allow">
<Conditions>
<FilePathCondition Path="*" />
</Conditions>
</FilePathRule>
<FilePathRule Id="dfaef909-a8b8-4be3-baf3-a1994a9845ab" Name="All files allowed, except specific Microsoft Office apps" Description="Configured for all users." UserOrGroupSid="S-1-1-0" Action="Allow">
<Conditions>
<FilePathCondition Path="*" />
</Conditions>
<Exceptions>
<FilePathCondition Path="*excel.exe" />
<FilePathCondition Path="*outlook.exe" />
<FilePathCondition Path="*powerpnt.exe" />
<FilePathCondition Path="*winword.exe" />
<FilePathCondition Path="*ONENOTE.EXE" />
<FilePathCondition Path="*MSPUB.EXE" />
<FilePathCondition Path="*MSACCESS.EXE" />
</Exceptions>
</FilePathRule>
</RuleCollection>
"@)

New-CimInstance -Namespace $namespaceName -ClassName $className -Property @{ParentID=$parentID;InstanceID="EXE";Policy=$obj}
