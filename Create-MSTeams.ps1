<#
        .SYNOPSIS
        Creates a Microsoft Team from parameters and outputs the sharepoint document library url.

        .DESCRIPTION
        Script is used to create an Microsoft Team with parameters and output the sharepoint document libary url that is being used for mapping it to client devices.

        .PARAMETER Name
        -TeamName
            Name of the Microsoft Team to be created.
            
            Required?                    true
            Default value
            Accept pipeline input?       false
            Accept wildcard characters?  false   

        -TeamDescription
            Description for the Microsoft Team.
            
            Required?                    true
            Default value
            Accept pipeline input?       false
            Accept wildcard characters?  false

        -TeamOwner
            Owner's email for the team.
            
            Required?                    true
            Default value
            Accept pipeline input?       false
            Accept wildcard characters?  false

        -TeamAlias
            Alias for the team.
            
            Required?                    true
            Default value
            Accept pipeline input?       false
            Accept wildcard characters?  false

        -TenantDomain
            Microsoft 365 tenant domain (without https://)
            
            Required?                    true
            Default value
            Accept pipeline input?       false
            Accept wildcard characters?  false

        .PARAMETER Extension

        .EXAMPLE
        C:\PS> Create-MSTeams.ps1 -TeamName "Finance Department" -TeamDescription "Finance Department documents" -TeamOwner peter@contoso.com -TeamAlias "FinanceDepartment" -TenantDomain contoso

        .COPYRIGHT
        MIT License, feel free to distribute and use as you like, please leave author information.

       .LINK
        BLOG: http://www.apento.com
        Twitter: @dk_hcandersen

        .DISCLAIMER
        This script is provided AS-IS, with no warranty - Use at own risk.
    #>


param (
    [Parameter(Mandatory=$true)]
    [string]$TeamName,            

    [Parameter(Mandatory=$true)]
    [string]$TeamDescription,     

    [Parameter(Mandatory=$true)]
    [string]$TeamOwner,           
    [Parameter(Mandatory=$true)]
    [string]$TeamAlias,           

    [Parameter(Mandatory=$true)]
    [string]$TenantDomain,        

    [Parameter(Mandatory=$false)]
    [int]$RetryCount = 5,         # Number of times to retry for SharePoint URL availability

    [Parameter(Mandatory=$false)]
    [int]$RetryInterval = 10      # Time in seconds to wait between retries
)

# Function to check and install required modules
function Install-ModuleIfMissing {
    param (
        [string]$ModuleName
    )
    
    if (-not (Get-Module -ListAvailable -Name $ModuleName)) {
        Write-Host "Module $ModuleName is missing. Installing..."
        Install-Module -Name $ModuleName -Force -AllowClobber
    } else {
        Write-Host "Module $ModuleName is already installed."
    }
}

# Check and install the required modules
Install-ModuleIfMissing -ModuleName "MicrosoftTeams"
Install-ModuleIfMissing -ModuleName "ExchangeOnlineManagement"

# Load required modules
Import-Module MicrosoftTeams
Import-Module ExchangeOnlineManagement

# Authenticate to Microsoft Teams
Connect-MicrosoftTeams

# Authenticate to Exchange Online
Connect-ExchangeOnline

# Check if a team with the same name or alias already exists
$existingTeam = Get-Team | Where-Object { $_.DisplayName -eq $TeamName -or $_.MailNickName -eq $TeamAlias }

if ($existingTeam) {
    Write-Host "A team with the name '$TeamName' or alias '$TeamAlias' already exists."
} else {
    # Create a new Microsoft 365 Group (the foundation for a Team)
    $newGroup = New-Team -DisplayName $TeamName -Description $TeamDescription -MailNickName $TeamAlias -Owner $TeamOwner

    if ($newGroup -ne $null) {
        Write-Host "Team created successfully!"

        # Initialize retry variables
        $attempts = 0
        $sharePointUrl = $null
        $documentLibraryUrl = $null

        # Loop to retry retrieving the SharePoint URL
        do {
            # Increment the retry attempt counter
            $attempts++
            
            # Retrieve the SharePoint URL using the Get-UnifiedGroup cmdlet
            try {
                $unifiedGroup = Get-UnifiedGroup -Identity $newGroup.GroupId
                
                if ($unifiedGroup -ne $null) {
                    $sharePointUrl = $unifiedGroup.SharePointDocumentsUrl

                    if ($sharePointUrl -ne $null -and $sharePointUrl -ne "") {
                        # Construct the SharePoint document library URL
                        $documentLibraryUrl = "$sharePointUrl"
                        Write-Host "SharePoint Document Library URL: $documentLibraryUrl"
                    }
                }
            } catch {
                Write-Host "Failed to retrieve the SharePoint URL. Attempt $attempts of $RetryCount."
            }
            
            # If no URL is retrieved, wait for the retry interval
            if ($sharePointUrl -eq $null) {
                Start-Sleep -Seconds $RetryInterval
            }

        } while ($sharePointUrl -eq $null -and $attempts -lt $RetryCount)

        # Check if the URL was successfully retrieved or max retries reached
        if ($sharePointUrl -eq $null) {
            Write-Host "Unable to retrieve the SharePoint document library URL after $RetryCount attempts."
        }
    }
    else {
        Write-Host "Failed to create the Team."
    }
}

# Disconnect from the services
Disconnect-MicrosoftTeams
Disconnect-ExchangeOnline -Confirm:$false
