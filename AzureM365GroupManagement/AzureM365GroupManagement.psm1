# https://docs.microsoft.com/microsoft-365/admin/create-groups/manage-creation-of-groups

function Set-M365GroupCreationAllowedGroup {

    <#
        .SYNOPSIS
        Configures the allowed group that can create Microsoft 365 Groups.

        .DESCRIPTION
        Configures the allowed group that can create Microsoft 365 Groups. Groups are identified by ObjectId or Name.

        .PARAMETER DisplayName
        The full name of the group in Azure Active Directory. (Required)

        .PARAMETER ObjectId
        The Group or ObjectId of the group in Azure Active Directory (e.g. fd4ec70a-274a-4c23-9c47-5dbc1a69c342). (Required)

        .EXAMPLE
        Set-M365GroupCreationAllowed Group -DisplayName "Allowed M365 Group Creators"

        This example uses the name of the group to configure the allowed group setting.

        .EXAMPLE
        Set-M365GroupCreationAllowed Group -ObjectId fd4ec70a-274a-4c23-9c47-5dbc1a69c342

        This example uses the group or objectId of the group to configure the allowed group setting.
    #>

    [CmdletBinding(
        DefaultParameterSetName = 'DisplayName',
        SupportsShouldProcess,
        ConfirmImpact = 'Medium'
    )]

    param(
        [Parameter(Mandatory, ParameterSetName = 'DisplayName')]
        [string]
        $DisplayName,
        [Parameter(Mandatory, ParameterSetName = 'ObjectId')]
        [string]
        $ObjectId
    )

    if (!(Test-GroupUnifiedDirectorySetting)) {
        Write-Warning -Message "No Group.Unified Directory Setting currently exists. Run New-GroupUnifiedDirectorySetting to create Group.Unified directory setting first."
        RETURN
    }

    if ($PSBoundParameters.ContainsKey("DisplayName")) {
        $groupFound = Get-AzureADGroup -SearchString $DisplayName

        switch ($groupFound.Count) {
            0 { Write-Error -Message "No Azure AD groups match the name $DisplayName. Please try again."; RETURN }
            1 { $groupFoundId = $groupFound.ObjectId; break }
            2 { Write-Error -Message "Multiple Azure AD Groups found matching $DisplayName. Please try again."; RETURN }
            Default { Write-Warning -Message "Something else went wrong with $DisplayName."; RETURN }
        }
    }

    if ($PSBoundParameters.ContainsKey("ObjectId")) {
        try {
            $groupFound = Get-AzureADGroup -ObjectId $ObjectId -ErrorAction STOP
        }
        catch {
            Write-Error -Message "Unable to find a group matching $ObjectId"
            RETURN
        }
        $groupFoundId = $groupFound.ObjectId
    }

    $groupUnifiedObject = Get-AzureADDirectorySetting | Where-Object -Property DisplayName -EQ -Value "Group.Unified"
    $groupUnifiedObject["GroupCreationAllowedGroupId"] = $groupFoundId

    try {
        if ($PSCmdlet.ShouldProcess($groupFoundId)) {
            Set-AzureADDirectorySetting -Id $groupUnifiedObject.Id -DirectorySetting $groupUnifiedObject -ErrorAction STOP
            Get-GroupUnifiedDirectorySettings
        }
    }
    catch {
        Write-Error -Message "Error enabling Group.Unified Azure AD Directory Setting: $($_.Exception)"
        RETURN
    }
} # End of Set-M365GroupCreationAllowedGroup

function Remove-M365GroupCreationAllowedGroup {

    <#
        .SYNOPSIS
        Clears the group setting for the group that is allowed to create Microsoft 365 Groups.

        .DESCRIPTION
        Clears the group setting for the group that is allowed to create Microsoft 365 Groups.

        .EXAMPLE
        Remove-M365GroupCreationAllowedGroup

        This example clears the allowed group setting for any group configured to create Microsoft 365 Groups.
    #>

    [CmdletBinding(
        SupportsShouldProcess,
        ConfirmImpact = 'Medium'
    )]
    param()

    if (Test-GroupUnifiedDirectorySetting) {
        $groupUnifiedObject = Get-AzureADDirectorySetting | Where-Object -Property DisplayName -EQ -Value "Group.Unified"
        $currentGroupId = $groupUnifiedObject.Values | Where-Object -Property 'Name' -EQ 'GroupCreationAllowedGroupId' | Select-Object -ExpandProperty Value
        $groupUnifiedObject["GroupCreationAllowedGroupId"] = ""

        try {
            if ($PSCmdlet.ShouldProcess($currentGroupId, 'Clearing GroupCreationAllowedGroupId')) {
                Set-AzureADDirectorySetting -Id $groupUnifiedObject.Id -DirectorySetting $groupUnifiedObject -ErrorAction STOP
                Get-GroupUnifiedDirectorySettings
            }
        }
        catch {
            Write-Error -Message "Error clearing GroupCreationAllowedGroupId Azure AD Directory Setting: $($_.Exception)"
            RETURN
        }
    }
    else {
        Write-Warning -Message "No Group.Unified Directory Setting currently exists. No changes being made."
    }
} # End of Remove-M365GroupCreationAllowedGroup

function Enable-M365GroupCreation {

    <#
        .SYNOPSIS
        Configures Microsoft 365 Group creation to True.

        .DESCRIPTION
        Configures Microsoft 365 Group creation to True. This allows all users in the tenant to create Microsoft 365 Groups.

        .EXAMPLE
        Enable-M365GroupCreation

        This example configures Microsoft 365 Group creation to True.
    #>

    [CmdletBinding(
        SupportsShouldProcess,
        ConfirmImpact = 'Medium'
    )]
    param()

    if (Test-GroupUnifiedDirectorySetting) {
        $groupUnifiedObject = Get-AzureADDirectorySetting | Where-Object -Property DisplayName -Value "Group.Unified" -EQ
        $groupUnifiedObject["EnableGroupCreation"] = "True"

        try {
            if ($PSCmdlet.ShouldProcess('EnableGroupCreation', 'Setting value to True')) {
                Set-AzureADDirectorySetting -Id $groupUnifiedObject.Id -DirectorySetting $groupUnifiedObject -ErrorAction STOP
                Get-GroupUnifiedDirectorySettings
            }
        }
        catch {
            Write-Error -Message "Error enabling Group.Unified Azure AD Directory Setting: $($_.Exception)"
            RETURN
        }
    }
    else {
        Write-Warning -Message "No Group.Unified Directory Setting currently exists. Run New-GroupUnifiedDirectorySetting to create Group.Unified directory setting first."
        RETURN
    }
} # End of Enable-M365GroupCreation

function Disable-M365GroupCreation {

    <#
        .SYNOPSIS
        Configures Microsoft 365 Group creation to False.

        .DESCRIPTION
        Configures Microsoft 365 Group creation to False. This prevents users in the tenant from creating Microsoft 365 Groups.

        .EXAMPLE
        Disable-M365GroupCreation

        This example configures Microsoft 365 Group creation to False.
    #>

    [CmdletBinding(
        SupportsShouldProcess,
        ConfirmImpact = 'Medium'
    )]
    param()

    if (Test-GroupUnifiedDirectorySetting) {
        $groupUnifiedObject = Get-AzureADDirectorySetting | Where-Object -Property DisplayName -Value "Group.Unified" -EQ
        $groupUnifiedObject["EnableGroupCreation"] = "False"
        try {
            if ($PSCmdlet.ShouldProcess('EnableGroupCreation', 'Setting value to False')) {
                Set-AzureADDirectorySetting -Id $groupUnifiedObject.Id -DirectorySetting $groupUnifiedObject -ErrorAction STOP
                Get-GroupUnifiedDirectorySettings
            }
        }
        catch {
            Write-Error -Message "Error enabling Group.Unified Azure AD Directory Setting: $($_.Exception)"
            RETURN
        }
    }
    else {
        Write-Warning -Message "No Group.Unified Directory Setting currently exists. No changes being made."
    }
} # End of Disable-M365GroupCreation

function Set-M365GroupUsageGuidelinesUrl {

    <#
        .SYNOPSIS
        Configures the URL for the Group Usage Guidelines.

        .DESCRIPTION
        Configures the URL for the Group Usage Guidelines.

        .PARAMETER URL
        The URL to configure as the Group Usage Guidelines. Should be a properly formatted HTTP URL. (Required)

        .EXAMPLE
        Set-M365GroupUsageGuidelinesUrl -URL "https://guidelines.jeffbrown.tech"

        This will set the Group Usage Guidelines URL to http://guidelines.jeffbrown.tech.
    #>

    [CmdletBinding(
        SupportsShouldProcess,
        ConfirmImpact = 'Low'
    )]
    param (
        [Parameter(Mandatory)]
        [string]
        $URL
    )

    if (Test-GroupUnifiedDirectorySetting) {
        $groupUnifiedObject = Get-AzureADDirectorySetting | Where-Object -Property DisplayName -EQ -Value "Group.Unified"
        $groupUnifiedObject["UsageGuidelinesUrl"] = $URL

        try {
            if ($PSCmdlet.ShouldProcess('UsageGuidelinesUrl', "Configuring $URL")) {
                Set-AzureADDirectorySetting -Id $groupUnifiedObject.Id -DirectorySetting $groupUnifiedObject -ErrorAction STOP
                Get-GroupUnifiedDirectorySettings
            }
        }
        catch {
            Write-Error -Message "Error enabling Group.Unified Azure AD Directory Setting: $($_.Exception)"
            RETURN
        }
    }
    else {
        Write-Warning -Message "No Group.Unified Directory Setting currently exists. Run New-GroupUnifiedDirectorySetting to create Group.Unified directory setting first."
    }
} # End of Set-M365GroupUsageGuidelinesUrl

function Remove-M365GroupUsageGuidelinesUrl {

    <#
        .SYNOPSIS
        Removes the URL for the Group Usage Guidelines.

        .DESCRIPTION
        Removes the URL for the Group Usage Guidelines.

        .EXAMPLE
        Remove-M365GroupUsageGuidelinesUrl

        This example removes the URL for the Group Usage Guidelines.
    #>

    [CmdletBinding(
        SupportsShouldProcess,
        ConfirmImpact = 'Medium'
    )]
    param()

    if (Test-GroupUnifiedDirectorySetting) {
        $groupUnifiedObject = Get-AzureADDirectorySetting | Where-Object -Property DisplayName -EQ -Value "Group.Unified"
        $currentURL = $groupUnifiedObject["UsageGuidelinesUrl"]
        $groupUnifiedObject["UsageGuidelinesUrl"] = ""
        try {
            if ($PSCmdlet.ShouldProcess('UsageGuidelinesUrl', "Removing $currentURL")) {
                Set-AzureADDirectorySetting -Id $groupUnifiedObject.Id -DirectorySetting $groupUnifiedObject -ErrorAction STOP
                Get-GroupUnifiedDirectorySettings
            }
        }
        catch {
            Write-Error -Message "Error enabling Group.Unified Azure AD Directory Setting: $($_.Exception)"
            RETURN
        }
    }
    else {
        Write-Warning -Message "No Group.Unified Directory Setting currently exists. No changes being made."
    }
} # End of Remove-M365GroupUsageGuidelinesUrl

function Add-CustomBlockedWord {

    <#
        .SYNOPSIS
        Adds a word to the CustomBlockedWordsList for Microsoft 365 groups.

        .DESCRIPTION
        Adds a word to the CustomBlockedWordsList for Microsoft 365 groups.

        .PARAMETER Word
        The word to add to the CustomBlockedWordsList property. (Required)

        .EXAMPLE
        Add-CustomBlockedWord -Word 'CEO'

        This example adds the word 'CEO' to the CustomBlockedWordsList property.

        .EXAMPLE
        Add-CustomBlockedWord -Word 'CEO','HR'

        This example adds the words 'CEO' and 'HR' to the CustomBlockedWordsList property.
    #>

    [CmdletBinding(
        SupportsShouldProcess,
        ConfirmImpact = 'Low'
    )]
    param (
        [Parameter(Position = 0, Mandatory)]
        [string[]]
        $Word
    )

    begin {
        $groupUnifiedObject = Get-GroupUnifiedDirectoryObject
        $blockedWordList = $groupUnifiedObject["CustomBlockedWordsList"]
        [System.Collections.ArrayList]$blockedWordArray = $blockedWordList.Split(",")
        $groupUnifiedId = ($groupUnifiedObject).Id
        $wordListModified = $false
    }

    process {
        foreach ($item in $Word) {

            if (Find-BlockedWord -Word $item -ArrayToSearch $blockedWordArray) {
                Write-Warning -Message "$item is already listed in the custom blocked word list"
            }
            else {
                if ($PSCmdlet.ShouldProcess('CustomBlockedWordsList', "Adding $item")) {
                    $blockedWordArray.Add($item) | Out-Null
                    $wordListModified = $true
                }
            }
        }
    }

    end {
        if ($wordListModified) {
            $groupUnifiedObject["CustomBlockedWordsList"] = $blockedWordArray -join ","

            try {
                if ($PSCmdlet.ShouldProcess('CustomBlockedWordsList')) {
                    Set-AzureADDirectorySetting -Id $groupUnifiedId -DirectorySetting $groupUnifiedObject -ErrorAction STOP
                    Get-GroupUnifiedDirectorySettings | Where-Object -Property Name -EQ -Value 'CustomBlockedWordsList'
                }
            }
            catch {
                Write-Error -Message "Error setting Group.Unified Blocked Words List: $($_.Exception)"
                RETURN
            }
        }
    }
} # End of Add-CustomBlockedWord

function Remove-CustomBlockedWord {

    <#
        .SYNOPSIS
        Removes a word from the CustomBlockedWordsList for Microsoft 365 groups.

        .DESCRIPTION
        Removes a word from the CustomBlockedWordsList for Microsoft 365 groups.

        .PARAMETER Word
        The word to remove from the CustomBlockedWordsList property. (Required)

        .EXAMPLE
        Remove-CustomBlockedWord -Word 'CEO'

        This example removes the word 'CEO' from the CustomBlockedWordsList property.

        .EXAMPLE
        remove-CustomBlockedWord -Word 'CEO','HR'

        This example removes the words 'CEO' and 'HR' from the CustomBlockedWordsList property.
    #>

    [CmdletBinding(
        SupportsShouldProcess,
        ConfirmImpact = 'Low'
    )]
    param (
        [Parameter(Position = 0, Mandatory)]
        [string[]]
        $Word
    )

    begin {
        $groupUnifiedObject = Get-GroupUnifiedDirectoryObject
        $blockedWordList = $groupUnifiedObject["CustomBlockedWordsList"]
        [System.Collections.ArrayList]$blockedWordArray = $blockedWordList.Split(",")
        $groupUnifiedId = ($groupUnifiedObject).Id
        $wordListModified = $false
    }

    process {
        foreach ($item in $Word) {
            if (Find-BlockedWord -Word $item -ArrayToSearch $blockedWordArray) {
                if ($PSCmdlet.ShouldProcess('CustomBlockedWordsList', "Adding $item")) {
                    $blockedWordArray.Remove($item) | Out-Null
                    $wordListModified = $true
                }
            }
            else {
                Write-Warning -Message "$item is not listed in the custom blocked word list for removal"
            }
        }
    }

    end {
        if ($wordListModified) {
            $groupUnifiedObject["CustomBlockedWordsList"] = $blockedWordArray -join ","

            try {
                if ($PSCmdlet.ShouldProcess('CustomBlockedWordsList')) {
                    Set-AzureADDirectorySetting -Id $groupUnifiedId -DirectorySetting $groupUnifiedObject -ErrorAction STOP
                    Get-GroupUnifiedDirectorySettings | Where-Object -Property Name -EQ -Value 'CustomBlockedWordsList'
                }
            }
            catch {
                Write-Error -Message "Error setting Group.Unified Blocked Words List: $($_.Exception)"
                RETURN
            }
        }
    }
} # End of Remove-CustomBlockedWord

function Get-CustomBlockedWordsList {
    <#
        .SYNOPSIS
        Lists the value of CustomBlockedWordsList property.

        .DESCRIPTION
        Lists the value of CustomBlockedWordsList property.

        .OUTPUTS
        System.String : List of blocked words separated by commas

        .EXAMPLE
        Get-CustomBlockedWordsList

        Output: CEO,HR,Executive
    #>

    [CmdletBinding()]
    [OutputType('System.String')]
    param ()

    Get-GroupUnifiedDirectorySettings | Where-Object -Property Name -EQ -Value 'CustomBlockedWordsList' | Select-Object -ExpandProperty Value
}

function Find-BlockedWord {
    <#
        .SYNOPSIS
        Finds if a blocked word exists in an array.

        .DESCRIPTION
        Finds if a blocked word exists in an array.
        This is an internal function to the module and should not be exported.

        .PARAMETER Word
        The word to search for in the array.

        .PARAMETER ArrayToSearch
        The array to search for the matching word.

        .OUTPUTS
        System.Boolean : Returns true or false if the word exists in the array

        .EXAMPLE
        Find-BlockedWord -Word 'CEO' -ArrayToSearch $currentWordArray

        This example will search $currentWordArray for the string 'CEO' and return true or false if it does.
    #>

    [CmdletBinding()]
    [OutputType('System.Boolean')]
    param (
        [Parameter(Mandatory)]
        [string]
        $Word,

        [Parameter(Mandatory)]
        [array]
        $ArrayToSearch
    )

    [bool]$found = $false

    if ($ArrayToSearch -icontains $Word) { $found = $true }

    $found
}

function Test-GroupUnifiedDirectorySetting {

    <#
        .SYNOPSIS
        Tests for the existence of a Group.Unified Directory Setting.

        .DESCRIPTION
        Tests for the existence of a Group.Unified Directory Setting.
        This is an internal function to the module and should not be exported.

        .OUTPUTS
        System.Boolean

        .EXAMPLE
        Test-GroupUnifiedDirectorySetting

        This example tests for the existence of a Group.Unified Directory Setting and returns $true or $false.
    #>

    [CmdletBinding()]
    [OutputType('System.Boolean')]
    param ()

    $foundGroupUnified = Get-AzureADDirectorySetting | Where-Object -Property DisplayName -EQ -Value "Group.Unified"
    if ($null -eq $foundGroupUnified) { RETURN $false } else { RETURN $true }
} # End of Test-GroupUnifiedDirectorySetting

function Get-GroupUnifiedDirectoryObject {
    <##>

    [CmdletBinding()]
    param ()

    $groupUnifiedObject = Get-AzureADDirectorySetting | Where-Object -Property DisplayName -EQ -Value "Group.Unified"
    $groupUnifiedObject
}
function Get-GroupUnifiedDirectorySettings {

    <#
        .SYNOPSIS
        Displays the current Group.Unified Directory Settings.

        .DESCRIPTION
        Displays the current Group.Unified Directory Settings.

        .EXAMPLE
        Get-GroupUnifiedDirectorySettings

        This example displays the current Group.Unified Directory Settings
    #>

    [CmdletBinding()]
    param ()

    try {
        $groupUnifiedObject = (Get-AzureADDirectorySetting | Where-Object -Property DisplayName -EQ -Value "Group.Unified" -ErrorAction STOP).Values
        if ($null -eq $groupUnifiedObject) {
            Write-Warning -Message "No Group.Unified Directory Setting currently exists. Run New-GroupUnifiedDirectorySetting to create Group.Unified directory setting first."
            RETURN
        }

        $groupUnifiedObject
    }
    catch {
        Write-Error -Message "Error getting Group.Unified Azure AD Directory Setting: $($_.Exception)"
    }
} # End of Get-GroupUnifiedDirectorySettings

function New-GroupUnifiedDirectorySetting {

    <#
        .SYNOPSIS
        Creates a new Azure AD Directory Setting using the Group.Unified template.

        .DESCRIPTION
        Creates a new Azure AD Directory Setting using the Group.Unified template.

        .EXAMPLE
        New-GroupUnifiedDirectorySetting

        Creates a new Azure AD Directory Setting using the Group.Unified template.
    #>

    [CmdletBinding(
        SupportsShouldProcess,
        ConfirmImpact = 'Medium'
    )]
    param()

    if (Test-GroupUnifiedDirectorySetting) {
        Write-Warning -Message "Group.Unified directory setting already exists."
        RETURN
    }
    else {
        try {
            Write-Verbose -Message "Creating new Azure AD Directory Setting using Group.Unified template"
            if ($PSCmdlet.ShouldProcess('Group.Unified', 'Creating new directory setting')) {
                $template = Get-AzureADDirectorySettingTemplate | Where-Object -Propert DisplayName -EQ -Value "Group.Unified"
                $newDirectorySetting = $template.CreateDirectorySetting()
                New-AzureADDirectorySetting -DirectorySetting $newDirectorySetting
            }
        }
        catch {
            Write-Error -Message "Error creating Group.Unified Azure AD Directory Setting: $($_.Exception)"
        }
    }
} # End of New-GroupUnifiedDirectorySetting

function Remove-GroupUnifiedDirectorySetting {

    <#
        .SYNOPSIS
        Removes the Group.Unified Directory Setting in Azure AD.

        .DESCRIPTION
        Removes the Group.Unified Directory Setting in Azure AD. This will remove any control or settings around Microsoft 365 Groups.

        .EXAMPLE
        Remove-GroupUnifiedDirectorySetting

        This example will remove the Group.Unified Directory Setting in Azure AD.
    #>

    [CmdletBinding(
        SupportsShouldProcess,
        ConfirmImpact = 'High'
    )]
    param()

    if (Test-GroupUnifiedDirectorySetting) {
        $groupUnifiedObject = Get-AzureADDirectorySetting | Where-Object -Property DisplayName -EQ -Value "Group.Unified"

        if ($PSCmdlet.ShouldProcess('Group.Unified', 'Removing existing directory setting')) {
            Remove-AzureADDirectorySetting -Id $groupUnifiedObject.Id
        }
    }
    else {
        Write-Warning -Message "No Group.Unified Directory Setting currently exists. No changes being made."
    }
} # End of Remove-GroupUnifiedDirectorySetting