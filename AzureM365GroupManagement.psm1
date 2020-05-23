# https://docs.microsoft.com/microsoft-365/admin/create-groups/manage-creation-of-groups

function Set-M365GroupCreationAllowedGroup {
    [CmdletBinding(DefaultParameterSetName = 'GroupName')]
    param(
        [Parameter(Mandatory, ParameterSetName = 'GroupName')]
        [string]
        $GroupName,
        [Parameter(Mandatory, ParameterSetName = 'GroupId')]
        [string]
        $GroupId
    )
    if ((Test-GroupUnifiedDirectorySetting) -eq $false) {
        Write-Warning -Message "No Group.Unified Directing Setting currently exists. Run New-GroupUnifiedDirectorySetting to create Group.Unified directory setting first."
        RETURN
    }
    if ($PSBoundParameters.ContainsKey("GroupName")) {
        $groupFound = Get-AzureADGroup -SearchString $GroupName
            
        switch ($groupFound.Count) {
            0 { Write-Error -Message "No Azure AD groups match the name $GroupName. Please try again."; RETURN }
            1 {
                $groupFoundId = $groupFound.ObjectId
                break
            }
            2 { Write-Error -Message "Multiple Azure AD Groups matching $GroupName. Please try again."; RETURN }
            Default { Write-Warning -Message "Something else went wrong with $GroupName."; RETURN }
        }
    }
    if ($PSBoundParameters.ContainsKey("GroupId")) {
        try {
            $groupFound = Get-AzureADGroup -ObjectId $GroupId -ErrorAction STOP
        }
        catch {
            Write-Error -Message "Unable to find a group matching $GroupId"
            RETURN
        }
        $groupFoundId = $groupFound.ObjectId
    }
    $groupUnifiedObject = Get-AzureADDirectorySetting | Where-Object -Property DisplayName -EQ -Value "Group.Unified"
    $groupUnifiedObject["GroupCreationAllowedGroupId"] = $groupFoundId
        
    try {
        Set-AzureADDirectorySetting -Id $groupUnifiedObject.Id -DirectorySetting $groupUnifiedObject -ErrorAction STOP
        Get-AzureADDirectorySetting -Id $groupUnifiedObject.Id | Select-Object -ExpandProperty Values
    }
    catch {
        Write-Error -Message "Error enabling Group.Unified Azure AD Directory Setting: $($Error[0])"
        RETURN
    }
} # End of Set-M365GroupCreationAllowedGroup

function Remove-M365GroupCreationAllowedGroup {
    [CmdletBinding()]
    param ()
    if ((Test-GroupUnifiedDirectorySetting) -eq $false) {
        Write-Warning -Message "No Group.Unified Directing Setting currently exists. No changes being made."
    }
    else {
        $groupUnifiedObject = Get-AzureADDirectorySetting | Where-Object -Property DisplayName -Value "Group.Unified" -EQ
        $groupUnifiedObject["GroupCreationAllowedGroupId"] = ""
        try {
            Set-AzureADDirectorySetting -Id $groupUnifiedObject.Id -DirectorySetting $groupUnifiedObject -ErrorAction STOP
            Get-AzureADDirectorySetting -Id $groupUnifiedObject.Id | Select-Object -ExpandProperty Values
        }
        catch {
            Write-Error -Message "Error enabling Group.Unified Azure AD Directory Setting: $($Error[0])"
            RETURN
        }
    }
} # End of Remove-M365GroupCreationAllowedGroup

function Enable-M365GroupCreation {
    [CmdletBinding()]
    param ()
    if ((Test-GroupUnifiedDirectorySetting) -eq $false) {
        Write-Warning -Message "No Group.Unified Directing Setting currently exists. Run New-GroupUnifiedDirectorySetting to create Group.Unified directory setting first."
        RETURN
    }
    else {
        $groupUnifiedObject = Get-AzureADDirectorySetting | Where-Object -Property DisplayName -Value "Group.Unified" -EQ
        $groupUnifiedObject["EnableGroupCreation"] = "True"
        
        try {
            Set-AzureADDirectorySetting -Id $groupUnifiedObject.Id -DirectorySetting $groupUnifiedObject -ErrorAction STOP
            Get-AzureADDirectorySetting -Id $groupUnifiedObject.Id | Select-Object -ExpandProperty Values
        }
        catch {
            Write-Error -Message "Error enabling Group.Unified Azure AD Directory Setting: $($Error[0])"
            RETURN
        }
    }
} # End of Enable-M365GroupCreation

function Disable-M365GroupCreation {
    [CmdletBinding()]
    param ()
    if ((Test-GroupUnifiedDirectorySetting) -eq $false) {
        Write-Warning -Message "No Group.Unified Directing Setting currently exists. No changes being made."
    }
    else {
        $groupUnifiedObject = Get-AzureADDirectorySetting | Where-Object -Property DisplayName -Value "Group.Unified" -EQ
        $groupUnifiedObject["EnableGroupCreation"] = "False"
        try {
            Set-AzureADDirectorySetting -Id $groupUnifiedObject.Id -DirectorySetting $groupUnifiedObject -ErrorAction STOP
            Get-AzureADDirectorySetting -Id $groupUnifiedObject.Id | Select-Object -ExpandProperty Values
        }
        catch {
            Write-Error -Message "Error enabling Group.Unified Azure AD Directory Setting: $($Error[0])"
            RETURN
        }
    }
} # End of Disable-M365GroupCreation

function Set-M365GroupUsageGuidelinesUrl {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]
        $URL
    )
    if ((Test-GroupUnifiedDirectorySetting) -eq $false) {
        Write-Warning -Message "No Group.Unified Directing Setting currently exists. Run New-GroupUnifiedDirectorySetting to create Group.Unified directory setting first."
    }
    else {
        $groupUnifiedObject = Get-AzureADDirectorySetting | Where-Object -Property DisplayName -Value "Group.Unified" -EQ
        $groupUnifiedObject["UsageGuidelinesUrl"] = $URL
        try {
            Set-AzureADDirectorySetting -Id $groupUnifiedObject.Id -DirectorySetting $groupUnifiedObject -ErrorAction STOP
            Get-AzureADDirectorySetting -Id $groupUnifiedObject.Id | Select-Object -ExpandProperty Values
        }
        catch {
            Write-Error -Message "Error enabling Group.Unified Azure AD Directory Setting: $($Error[0])"
            RETURN
        }
    }
} # End of Set-M365GroupUsageGuidelinesUrl

function Remove-M365GroupUsageGuidelinesUrl {
    [CmdletBinding()]
    param ()
    if ((Test-GroupUnifiedDirectorySetting) -eq $false) {
        Write-Warning -Message "No Group.Unified Directing Setting currently exists. No changes being made."
    }
    else {
        $groupUnifiedObject = Get-AzureADDirectorySetting | Where-Object -Property DisplayName -Value "Group.Unified" -EQ
        $groupUnifiedObject["UsageGuidelinesUrl"] = ""
        try {
            Set-AzureADDirectorySetting -Id $groupUnifiedObject.Id -DirectorySetting $groupUnifiedObject -ErrorAction STOP
            Get-AzureADDirectorySetting -Id $groupUnifiedObject.Id | Select-Object -ExpandProperty Values
        }
        catch {
            Write-Error -Message "Error enabling Group.Unified Azure AD Directory Setting: $($Error[0])"
            RETURN
        }
    }
} # End of Remove-M365GroupUsageGuidelinesUrl

function Test-GroupUnifiedDirectorySetting {
    [CmdletBinding()]
    param ()
    $foundGroupUnified = (Get-AzureADDirectorySetting | Where-Object -Property DisplayName -Value "Group.Unified" -EQ).Id
    if ($null -eq $foundGroupUnified) { RETURN $false } else { RETURN $true }    
} # End of Test-GroupUnifiedDirectorySetting

function New-GroupUnifiedDirectorySetting {
    [CmdletBinding()]
    param()
    Write-Verbose -Message "Creating new Azure AD Directory Setting using Group.Unified template"
    $template = Get-AzureADDirectorySettingTemplate | Where-Object -Propert DisplayName -Value "Group.Unified" -EQ
    $newDirectorySetting = $template.CreateDirectorySetting()
    New-AzureADDirectorySetting -DirectorySetting $newDirectorySetting
} # End of New-GroupUnifiedDirectorySetting

function Remove-GroupUnifiedDirectorySetting {
    [CmdletBinding()]
    param()
    if ((Test-GroupUnifiedDirectorySetting) -eq $false) {
        Write-Warning -Message "No Group.Unified Directing Setting currently exists. No changes being made."        
    }
    else {
        $groupUnifiedObject = Get-AzureADDirectorySetting | Where-Object -Propert DisplayName -Value "Group.Unified" -EQ
            
        $yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", "Removes Group.Unified directory setting"
        $no = New-Object System.Management.Automation.Host.ChoiceDescription "&No", "Discards any changes"
        $options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no)
        $result = $host.ui.PromptForChoice("Remove Azure AD Directory Setting", "Do you want to remove the Group.Unified directory setting with an ID of $($groupUnifiedObject.Id)?", $options, 0)
            
        switch ($result) {
            0 { Remove-AzureADDirectorySetting -Id $($groupUnifiedObject.Id); break }
            1 { Write-Output "No changes being made."; break }
        }
    }
} # End of Remove-GroupUnifiedDirectorySetting
    