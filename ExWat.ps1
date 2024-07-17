<#
.SYNOPSIS
    This is the tool for checking the expiration of the Azure objects
.DESCRIPTION
    ExpirationWatcher works using Graph API and Azure App Registration. More information can be found in Readme
.NOTES
    Version: 2.0.0
    Author: Aslan Imanalin
#>

#region Functions
function Write-Log {

    [CmdletBinding()]
    Param (
        
        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
        [ValidateNotNullOrEmpty()]
        [Alias("LogContent")]
        [string]$Message,
        
        [Parameter(Mandatory = $false)]
        [ValidateSet("Error", "Warn", "Info")]
        [string]$Level = "Info",
        
        [Parameter(Mandatory = $false)]
        [switch]$NoClobber
    
    )
    
    Begin {
        $FormattedDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    }
    Process {

        switch ($Level) {
            'Error' {
                $LevelText = 'ERROR:'
                Write-Error "$FormattedDate $LevelText $Message"
            }
            'Warn' {
                $LevelText = 'WARNING:'
                Write-Warning "$FormattedDate $LevelText $Message"
            }
            'Info' {
                $LevelText = 'INFO:'
                Write-Output "$FormattedDate $LevelText $Message"
            }
        }
        
    }
}

function Get-ConfigurationVariables {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $True)]
        [string]$RuntimeEnvironment,

        [string]$ResourceGroupName = $Env:EW_ResourceGroupName,

        [string]$AutomationAccountName = $Env:EW_AutomationAccountName
    )

    [array]$AutomationVariablesList = @(
        "EW_appId",
        "EW_tenantId",
        "EW_Secret",
        "EW_SenderAddress",
        "EW_TeamsWebHook",
        "EW_CCRecipientAddress",
        "EW_TableRG",
        "EW_TableName",
        "EW_TableSAName"
    )
    
    $Variables = @()

    switch ($RuntimeEnvironment) {
        "Cloud" {
            foreach ($Variable in $AutomationVariablesList) {
                try {
                    $VariableValue = Get-AutomationVariable -Name $Variable -ErrorAction "STOP"
                }
                catch {
                    throw "Error: $($_.Exception.Message)"
                    exit 1
                }
                    
                $Variables += [PSCustomObject]@{
                    name  = $Variable
                    value = $VariableValue
                }
            }
        }

        "Local" {
            $AutomationVariablesSplat = @{
                ResourceGroupName     = $ResourceGroupName
                AutomationAccountName = $AutomationAccountName
                ErrorAction           = "STOP"
            }
            try {
                $AutomationVariables = Get-AzAutomationVariable @AutomationVariablesSplat
            }
            catch {
                throw "Error: $($_.Exception.Message)"
                exit 1
            }

            $Variables = $AutomationVariables | where-object { $_.name -like "EW_*" }
            
            try {
                $ClientSecretSecureString = Get-Secret -Name $Env:EW_LocalSecretName -ErrorAction Stop
            }
            catch {
                throw "Error: $($_.Exception.Message)"
                exit 1
            }
                
            $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($ClientSecretSecureString)
            [string]$ClientSecret = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
            $Variables += [PSCustomObject]@{
                name  = "EW_Secret"
                value = $ClientSecret
            }


            try {
                $TeamsUrlSecureString = Get-Secret -Name $Env:EW_TeamsSecret -ErrorAction Stop
            }
            catch {
                throw "Error: $($_.Exception.Message)"
                exit 1
            }
                
            $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($TeamsUrlSecureString)
            [string]$TeamsWebHookUrl = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
            $Variables += [PSCustomObject]@{
                name  = "EW_TeamsWebHook"
                value = $TeamsWebHookUrl
            }
        }
    }
    $ConfigurationObject = [PSCustomObject]@{
        AppId              = $($Variables | Where-Object { $_.name -eq "EW_appId" }).value
        ListId             = $($Variables | Where-Object { $_.name -eq "EW_listId" }).value
        SiteId             = $($Variables | Where-Object { $_.name -eq "EW_siteId" }).value
        TenantID           = $($Variables | Where-Object { $_.name -eq "EW_tenantId" }).value
        SenderAddress      = $($Variables | Where-Object { $_.name -eq "EW_SenderAddress" }).value
        CCRecipientAddress = $($Variables | Where-Object { $_.name -eq "EW_CCRecipientAddress" }).value
        TeamsWebHookUrl    = [string]$($Variables | Where-Object { $_.name -eq "EW_TeamsWebHook" }).value
        Secret             = [string]$($Variables | Where-Object { $_.name -eq "EW_Secret" }).value
        TableRG            = [string]$($Variables | Where-Object { $_.name -eq "EW_TableRG" }).value
        TableName          = [string]$($Variables | Where-Object { $_.name -eq "EW_TableName" }).value
        TableSAName        = [string]$($Variables | Where-Object { $_.name -eq "EW_TableSAName" }).value
    }
    $ConfigurationObject
}

function Get-ExWatList {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $True)]
        [string]$ResourceGroup,

        [Parameter(Mandatory = $True)]
        [string]$StorageAccountName,

        [Parameter(Mandatory = $True)]
        [string]$TableName
    )
    
    [array]$List = @()

    $ObjectTypeFriendlyNameDict = [PSCustomObject]@{
        AzAppSecret           = "Azure AD Registered Application Secret"
        iOSLoBApp             = "iOS Line of Business Application"
        ApnCertificate        = "Apple Push Notification Certificate"
        AzEntAppKeyCredential = "Azure AD Enterprise Application SAML Certificate"
        AzAppCert             = "Azure AD Registered Application Certificate"
    }

    $StorageAccount = Get-AzStorageAccount -ResourceGroupName $ResourceGroup -Name $StorageAccountName

    $Context = $StorageAccount.Context

    $Table = (Get-AzStorageTable -Name $tableName -Context $Context).CloudTable

    $TableRows = Get-AzTableRow -Table $Table

    foreach ($Row in $TableRows) {
        if (!$Row.ChildObjectId) {
            $Row | Add-Member -MemberType NoteProperty -Name 'ChildObjectId' -Value 'n/a'
        }

        $List += [PSCustomObject]@{
            Name                   = $Row.title.ToString()
            Type                   = $Row.objectType.ToString()
            TypeFriendlyName       = $ObjectTypeFriendlyNameDict.$($Row.objectType.ToString())
            Id                     = $Row.objectId.ToString()
            ChildId                = $Row.childId.ToString()
            IsTestEnabled          = [System.Convert]::ToBoolean($Row.isEnabled)
            NotificationWindowDays = [int]$Row.notificationDays
            Owner                  = $Row.owner.ToString()
            NotificationRecipients = $Row.notificationRecipients -split ';'
        }
    }
    $List

}

function Get-ExpirationObject {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $True)]
        [string]
        $ExpirationDateString,

        [Parameter(Mandatory = $True)]
        [PSCustomObject]
        $CheckingObject
    )
    
    [datetime]$CurrentDate = Get-Date

    [datetime]$ExpirationDate = Get-Date $ExpirationDateString
    [datetime]$WarningDate = $ExpirationDate.AddDays(-$CheckingObject.NotificationWindowDays)
    [int]$DaysBeforeExpiration = $($ExpirationDate - $CurrentDate).Days
    [int]$WeeksBeforeExpiration = $DaysBeforeExpiration / 7
    if ($CurrentDate -ge $WarningDate) {
        [bool]$ExpirationWarning = $True
    }
    else {
        [bool]$ExpirationWarning = $False
    }
    if ($CurrentDate -gt $ExpirationDate) {
        [bool]$Expired = $True
    }
    else {
        [bool]$Expired = $False
    }

    $ObjectInfo = [PSCustomObject]@{
        Name                  = $CheckingObject.Name
        Id                    = $CheckingObject.Id
        ChildId               = $CheckingObject.ChildId
        IsTestEnabled         = $CheckingObject.IsTestEnabled
        OwnerName             = $CheckingObject.Owner
        OwnerEmail            = $CheckingObject.NotificationRecipients
        Type                  = $CheckingObject.Type
        TypeFriendlyName      = $CheckingObject.TypeFriendlyName
        ExpirationDate        = $ExpirationDate
        ExpirationWindowDays  = $CheckingObject.NotificationWindowDays
        DaysBeforeExpiration  = $DaysBeforeExpiration
        WeeksBeforeExpiration = $WeeksBeforeExpiration
        ExpirationWarning     = $ExpirationWarning
        Expired               = $Expired
        WarningDate           = $WarningDate
    }

    $ObjectInfo
}

function Get-ObjectWithChildExpiration {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $True)]
        [PSCustomObject]$CheckingObject,

        [Parameter(Mandatory = $True)]
        [string]$RequestResource,

        [Parameter(Mandatory = $True)]
        [PSCustomObject]$Token
    )
    
    try {
        $RequestSplat = @{
            Token       = $Token
            Resource    = $RequestResource
            ErrorAction = "Stop"
        }
        try {
            $RequestResponse = Invoke-GraphApiRequest @RequestSplat
        }
        catch {
            Write-Error "Error: $($_.Exception.Message)"
            exit 1
        }
    }
    catch {
        throw "Error: $($_.Exception.message)"
        exit 1
    }

    if (!($RequestResponse | Where-Object { $_.keyId -eq $CheckingObject.ChildId }).endDateTime) {
        throw "Error: $($CheckingObject.Name) does not have an expiration date."
        exit 1
    }

    [string]$ExpirationDateString = ($RequestResponse | Where-Object { $_.keyId -eq $CheckingObject.ChildId }).endDateTime
    $ObjectInfo = Get-ExpirationObject -ExpirationDateString $ExpirationDateString -CheckingObject $CheckingObject

    $ObjectInfo
}

function Get-ObjectExpiration {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $True)]
        [PSCustomObject]$CheckingObject,

        [Parameter(Mandatory = $True)]
        [string]$RequestResource,

        [Parameter(Mandatory = $True)]
        [PSCustomObject]$Token
    )

    try {
        $RequestSplat = @{
            Token       = $Token
            Resource    = $RequestResource
            ErrorAction = "Stop"
        }
        try {
            $RequestResponse = Invoke-GraphApiRequest @RequestSplat
        }
        catch {
            Write-Error "Error: $($_.Exception.Message)"
            exit 1
        }
    }
    catch {
        throw "Error: $($_.Exception.message)"
        exit 1
    }

    $ExpirationDateString = $RequestResponse.expirationDateTime

    $ObjectInfo = Get-ExpirationObject -ExpirationDateString $ExpirationDateString -CheckingObject $CheckingObject

    $ObjectInfo
    
}

function Send-ExpirationNotificationEmail {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [PSCustomObject[]]$Objects,

        [Parameter(Mandatory = $true)]
        [mailaddress]$SenderAddress,

        [Parameter(Mandatory = $true)]
        [mailaddress]$CCRecipientAddress,

        [Parameter(Mandatory = $True)]
        [PSCustomObject]$Token
    )
    
    foreach ($Object in $Objects) {
        Write-Log "Working with $($Object.Type) $($Object.Name) $($Object.Id) ($($Object.ChildId)) - $($Object.OwnerName)"
        if ($Object.DaysBeforeExpiration -gt 0) {
            $MailBody = "The <b>$($Object.TypeFriendlyName)  $($Object.Name) $($Object.Id) ($($Object.ChildId))</b> will expire in <b> $($Object.DaysBeforeExpiration)</b> days! <br>Expiration date: $($Object.ExpirationDate.toString("MM/dd/yyyy"))"
            $MailSubject = "Expiration Warning: $($Object.Type) - $($Object.Name)"
        }
        else {
            $MailBody = "The <b>$($Object.TypeFriendlyName)  $($Object.Name) $($Object.Id) ($($Object.ChildId))</b> expired on <b>$($Object.ExpirationDate.toString("MM/dd/yyyy"))</b>"
            $MailSubject = "Expiration Warning: $($Object.Type) - $($Object.Name)"
        }

        $MailRequestSplat = @{
            Token                  = $Token
            SenderUPN              = $SenderAddress
            Recipients             = $Object.OwnerEmail
            CopyRecipients         = $CCRecipientAddress
            MailSubject            = $MailSubject
            MessageBody            = $MailBody
            MessageBodyContentType = 'HTML'
            ErrorAction            = "STOP"
        }
        try {
            Send-GraphEmail @MailRequestSplat
            Write-Log "Message has been sent"
        }
        catch {
            throw "Message hasn't been sent. Error: $($_.Exception.Message)"
        }
    }
}

function Get-ObjectsSummary {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $True)]
        [PSCustomObject[]]$Objects
    )
    $ObjectsSummary = @()
    foreach ($Object in $Objects) {
        $ObjectsSummary += $Object | Select-Object Type, Name, Id, ChildId, ExpirationDate, DaysBeforeExpiration, ExpirationWarning
    }
    $ObjectsSummary    
}

function Get-TeamsTableSummary {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $True)]
        [PSCustomObject[]]$Objects
    )

    $FormattedSummary = $Objects | Select-Object @{
        L = "Object"; E = {
            if ($_.ExpirationWarning -eq $True -and $_.Expired -eq $False) {
                "⚠ *$($_.Type)*: $($_.Name)"
            }
            elseif ($_.Expired -eq $True) {
                "❌ *$($_.Type)*: $($_.Name)"
            }
            else {
                " ✔ *$($_.Type)*: $($_.Name)"
            }
        }
    }, @{L = "Days Before Expiration"; E = { $_.DaysBeforeExpiration } } | Sort-Object 'Days Before Expiration'
    
    $FormattedSummary | ConvertTo-Markdown  | Out-String

}

function Get-TeamsFormattedFacts {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $True)]
        [PSCustomObject[]]$Objects
    )

    $Summary = @()
    
    $Summary += New-TeamsFact -Name "Total Objects:" -Value $($Objects.Count)
    $Summary += New-TeamsFact -Name "Total Objects with Expiration Warning:" -Value $(
        $Objects | Where-Object { $_.ExpirationWarning -eq $True -and $_.Expired -eq $False } | Measure-Object | Select-Object -ExpandProperty Count
    )
    $Summary += New-TeamsFact -Name "Total Objects Expired:" -Value $(
        $Objects | Where-Object { $_.Expired -eq $True } | Measure-Object | Select-Object -ExpandProperty Count
    )

    $Summary
}
#endregion

#Requires -module GraphApiRequests, GraphEmailSender, PSTeams, PSMarkdown, AzTable

Import-Module GraphApiRequests, GraphEmailSender, PSTeams, PSMarkdown, AzTable

Write-Log "Starting the ExWat..."
$Environment = $PSPrivateMetadata.JobId
if ($Environment) {
    $RuntimeEnvironment = "Cloud"
}
else {
    $RuntimeEnvironment = "Local"
}
Write-Log "This is the $RuntimeEnvironment environment"

Write-Log "Getting the Automation configuration variables"
try {
    $AutomationConfiguration = Get-ConfigurationVariables -RuntimeEnvironment $RuntimeEnvironment -ErrorAction "Stop"
    Write-Log "Variables have been loaded"
}
catch {
    Write-Log -Level Error "Unable to get automation variables. $($_.Exception)"
    exit 1
}


Write-Log "Getting the Graph API Token"
$TokenSplat = @{
    AppId       = $AutomationConfiguration.AppId
    AppSecret   = $AutomationConfiguration.Secret
    TenantID    = $AutomationConfiguration.TenantID
    ErrorAction = "Stop"
}
try {
    $Token = Get-GraphToken @TokenSplat
    Write-Log "Token has been successfully issued"
}
catch {
    Write-Log -Level Error "Error: $($_.Exception.Message)"
    exit 1
}

try{
    Connect-AzAccount -identity -ErrorAction "Stop"
}
catch {
    Write-Log -Level Error "Unable to connect to Azure. $($_.Exception)"
    exit 1
}


Write-Log "Getting the Expiration List object"
try {
    $ExpirationListSplat = @{
        ResourceGroup      = $AutomationConfiguration.TableRG
        StorageAccountName = $AutomationConfiguration.TableSAName
        TableName          = $AutomationConfiguration.TableName
        ErrorAction        = "STOP"
    }
    $ExpirationListObject = Get-ExWatList @ExpirationListSplat
}
catch {
    Write-Log -Level Error "Unable to get ExWat List. $($_.Exception)"
    exit 1
}

$ExWatListToCheck = $ExpirationListObject | Where-object { $_.IsTestEnabled -eq "true" }

if ($ExWatListToCheck.Count -gt 0) {
    Write-Log "Overall objects: $($ExpirationListObject.Count)"
    Write-Log "Objects to check: $($ExWatListToCheck.Count)"
}
else {
    Write-Log -Level Warn "There is no object to check"
    exit 1
}

$CheckingObjectsList = @()

foreach ($Object in $ExWatListToCheck) {
    $ObjectTitle = "$($Object.Type) - $($Object.Name) $($Object.Id)($($Object.ChildId))"
    Write-Log "Working with $ObjectTitle"
    switch ($Object.Type) {
        "AzAppSecret" {
            $RequestResource = "applications/$($object.Id)/passwordCredentials"
            $AzAppSecretSplat = @{
                CheckingObject  = $Object
                RequestResource = $RequestResource
                Token           = $Token
                ErrorAction     = "Continue"
            }
            try {
                $Result = Get-ObjectWithChildExpiration @AzAppSecretSplat
            }
            catch {
                Write-Log -Level Error "Unable to get an expiration date for $ObjectTitle. $($_.Exception)"
                Write-Log -Level Warn "Going to the next one"
            }
        }

        "iOSLoBApp" {
            $RequestResource = "deviceAppManagement/mobileApps/$($Object.Id)"
            $iOSLoBAppSplat = @{
                CheckingObject  = $Object
                RequestResource = $RequestResource
                Token           = $Token
                ErrorAction     = "Continue"
            }
            try {
                $Result = Get-ObjectExpiration @iOSLoBAppSplat
            }
            catch {
                Write-Log -Level Error "Unable to get an expiration date for $ObjectTitle. $($_.Exception)"
                Write-Log -Level Warn "Going to the next one"
            }
        }

        "ApnCertificate" {
            $RequestResource = "deviceManagement/applePushNotificationCertificate"
            $ApnCertificateSplat = @{
                CheckingObject  = $Object
                RequestResource = $RequestResource
                Token           = $Token
                ErrorAction     = "Continue"
            }
            try {
                $Result = Get-ObjectExpiration @ApnCertificateSplat
            }
            catch {
                Write-Log -Level Error "Unable to get an expiration date for $ObjectTitle. $($_.Exception)"
                Write-Log -Level Warn "Going to the next one"
            }
        }

        "AzEntAppKeyCred" {
            $RequestResource = "servicePrincipals/$($Object.Id)/keyCredentials"
            $AzEntAppKeyCredSplat = @{
                CheckingObject  = $Object
                RequestResource = $RequestResource
                Token           = $Token
                ErrorAction     = "Continue"
            }
            try {
                $Result = Get-ObjectWithChildExpiration @AzEntAppKeyCredSplat
            }
            catch {
                Write-Log -Level Error "Unable to get an expiration date for $ObjectTitle. $($_.Exception)"
                Write-Log -Level Warn "Going to the next one"
            }
        }

        "AzAppCert" {
            $RequestResource = "applications/$($Object.Id)/keyCredentials"
            $AzAppCertSplat = @{
                CheckingObject  = $Object
                RequestResource = $RequestResource
                Token           = $Token
                ErrorAction     = "Continue"
            }
            try {
                $Result = Get-ObjectWithChildExpiration @AzAppCertSplat
            }
            catch {
                Write-Log -Level Error "Unable to get an expiration date for $ObjectTitle. $($_.Exception)"
                Write-Log -Level Warn "Going to the next one"
            }
        }
        
    }
    $CheckingObjectsList += $Result
    $ExpirationDate = $Result.ExpirationDate
    Write-Log "$($Object.ObjectType) - $($Object.ObjectName) expiration date: $ExpirationDate"
}


$ObjectsWithWarning = $CheckingObjectsList | Where-Object { $_.ExpirationWarning -eq $True }
$ObjectsWithWarningCount = $($ObjectsWithWarning | Measure-Object).Count
if ($ObjectsWithWarningCount -gt 0) {
    Write-Log "There objects that are going to expire or already expired count: $ObjectsWithWarningCount"
    Write-Log "Sending the expiration notifications"
    try {
        $MailSplat = @{
            Objects            = $ObjectsWithWarning
            SenderAddress      = $AutomationConfiguration.SenderAddress
            CCRecipientAddress = $AutomationConfiguration.CCRecipientAddress
            Token              = $Token
            ErrorAction        = 'Continue'
        }
        Send-ExpirationNotificationEmail @MailSplat
    }
    catch {
        Write-Log -Level Error "Unable to send notification. $($_.Exception)"
    }
}
else {
    Write-Log "There are no objects are going to expire or already expired"
}

$Summary = Get-ObjectsSummary  -Objects $CheckingObjectsList
Write-Log "Summary:"
$Summary | Format-Table -AutoSize -Wrap

$TeamsSummaryTable = Get-TeamsTableSummary -Objects $CheckingObjectsList
$FormattedFacts = Get-TeamsFormattedFacts -Objects $CheckingObjectsList

if ($ObjectsWithWarningCount -gt 0) {
    $TeamsMessageColor = "Red"
}
else {
    $TeamsMessageColor = "Green"
}

$SectionSplat = @{
    ActivityDetails = $FormattedFacts
    Text            = $TeamsSummaryTable
}
$Section = New-TeamsSection @SectionSplat

$TeamsSplat = @{
    Uri          = $AutomationConfiguration.TeamsWebHookUrl
    MessageTitle = "ExWat - Expiration Warnings | Summary"
    Color        = $TeamsMessageColor
    Sections     = $Section
    ErrorAction  = "STOP"
}

try {
    Send-TeamsMessage @TeamsSplat
}
catch {
    Write-Log -Level Error "Error: $($_.Exception.Message)"
    exit 1
}