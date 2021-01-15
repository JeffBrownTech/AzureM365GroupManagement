# AzureM365GroupManagement

This is a module for managing the Microsoft 365 Groups directory settings, such as configuring allowed creators for Microsoft 365 groups and other settings. This module currently uses commands available in the [AzureADPreview module](https://docs.microsoft.com/en-us/powershell/module/azuread/?view=azureadps-2.0-preview) available in the [PowerShell Gallery](https://docs.microsoft.com/en-us/powershell/module/azuread/?view=azureadps-2.0-preview).

![](https://jeffbrown.tech/wp-content/uploads/2021/01/Enable-M365GroupCreation.png)

*Disclaimer: Use of this code does not come with any support and is provided 'as-is'. Use at your own risk and review the code and test prior to use in production.*

# Submitting Issues
If you run into any issues or errors using this module, please submit an Issue here in GitHub and I will review. If you have an enhancement, submit an issue and label it as an enhancement. I will implement as time allows.

## Getting Started

This module requires an administrator account in Microsoft/Office 365 to administer Active Directory. The module also uses the AzureADPreview module, which contains the *Get-AzureADDirectorySetting* cmdlet. The module is available in the [PowerShell Gallery](https://docs.microsoft.com/en-us/powershell/module/azuread/?view=azureadps-2.0-preview), and that page contains instructions on how to install the module.

## Importing the Module
After downloading the files in this repository, you can import the module for use in a PowerShell console by importing the .PSD1 file:

```powershell
Import-Module AzureM365GroupManagement.psd1
```

You can then view the commands available in the module using the **Get-Command** cmdlet:

```powershell
Get-Command -Module AzureM365GroupManagement.psd1
```

The AzureM365GRoupManagement module is also available for download through the [PowerShell Gallery](https://www.powershellgallery.com/packages/AzureM365GroupManagement):

```powershell
Install-Module -Name AzureM365GroupManagement
```

## Overview of Group Creation

Many organizations want to control if their users can Microsoft 365 groups (formerly known as Office 365 groups). Typically organizations disable group creation to prevent users from creating teams in Microsoft Teams. The goal of this module to provide commands that can easily enable and disable the creation of Microsoft 365 groups. The module also provides commands to set the allowed group of users who can create groups.

For more information on managing who can create Microsoft 365 groups, check out this Microsoft Docs article [Manage who can create Microsoft 365 groups](https://docs.microsoft.com/en-us/microsoft-365/solutions/manage-creation-of-groups?view=o365-worldwide). The code available in the article is the inspiration for this module.

## Configuring Group Creation

Before configuring group creation settings, you first need to create the Group.Unified directory setting. You can verify if this directory setting already exists by using the **Get-GroupUnifiedDirectorySettings** module cmdlet. If directory setting does not exist, the command output directs you to run the **New-GroupUnifiedDirectorySetting** to create it.

```powershell
New-GroupUnifiedDirectorySetting
```

![](https://jeffbrown.tech/wp-content/uploads/2021/01/VerifyCreateGroupUnifiedSetting.png)

With the Group.Unified directory setting created, you can view the current settings by running the **Get-GroupUnifiedDirectorySettings** cmdlet.

```powershell
Get-GroupUnifiedDirectorySetting
```

![](https://jeffbrown.tech/wp-content/uploads/2021/01/Get-GroupUnifiedDirectorySettings.png)

Notice the value of *EnableGroupCreation* is currently set to "True". This means users in the organization are allowed to create Microsoft 365 groups in various services like Teams, Outlook, or Yammer. In order to disable group creation for the organization, use the **Disable-M365GroupCreation** cmdlet. Running this command will change *EnableGroupCreation* to "False".

```powershell
Disable-M365GroupCreation
```

![](https://jeffbrown.tech/wp-content/uploads/2021/01/Disable-M365GroupCreation.png)

Although group creation is now disabled, it is only disabled for regular users. Users in certain administrator roles still can, such as Global admins, Exchange, SharePoint, and Teams Administrators, and Directory Writers. If you want to designate a group of regular users who are not assigned to these roles, you can specify a group that is allowed. This value is stored in the *GroupCreationAllowedGroupId*.

If you don't have a security group already created to store these users, go ahead and create one now. Once the group is created, you can use either the group name or object ID with the **Set-M365GroupCreationAllowedGroup** cmdlet. An example of each is shown below.

```powershell
Set-M365GroupCreationAllowedGroup -DisplayName <Security Group Name>

Set-M365GroupCreationAllowedGroup -ObjectId <Security Group GUID>
```

![](https://jeffbrown.tech/wp-content/uploads/2021/01/SetAllowedGroup.png)

To remove the allowed group, use the **Remove-M365GroupCreationAllowedGroup** to clear the value.