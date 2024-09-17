<#
        .SYNOPSIS
        Creates a Microsoft Team from parameters and outputs the sharepoint document library url.

        .DESCRIPTION
        Script is used to create an Microsoft Team with parameters and output the sharepoint document libary url that is being used for mapping it to client devices.

        .PARAMETER Name
        -TeamName
            Name of the Microsoft Team to add users to.
            
            Required?                    true
            Default value
            Accept pipeline input?       false
            Accept wildcard characters?  false   

        -CsvPath
            Path to csv file with users to add to the Microsoft Team.
            
            Required?                    true
            Default value
            Accept pipeline input?       false
            Accept wildcard characters?  false

        .PARAMETER Extension

        .EXAMPLE
        C:\PS> Add-UsersToTeams.ps1 -TeamName "Finance Department" -CsvPath "C:\Contoso\users.csv"

        .COPYRIGHT
        MIT License, feel free to distribute and use as you like, please leave author information.

       .LINK
        BLOG: http://www.apento.com
        Twitter: @dk_hcandersen

        .DISCLAIMER
        This script is provided AS-IS, with no warranty - Use at own risk.
    #>



# Define the parameters
param (
    [Parameter(Mandatory=$true)]
    [string]$TeamName,            

    [Parameter(Mandatory=$true)]
    [string]$CsvPath 

)

# Function to check if a module is installed
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

# Ensure the MicrosoftTeams module is installed
Install-ModuleIfMissing -ModuleName "MicrosoftTeams"

# Import the MicrosoftTeams module
Import-Module MicrosoftTeams

# Connect to Microsoft Teams
Connect-MicrosoftTeams

# Get the Team ID based on the Team name
$team = Get-Team | Where-Object { $_.DisplayName -eq $teamName }

if ($null -eq $team) {
    Write-Host "Team '$teamName' not found!"
    Disconnect-MicrosoftTeams
    exit
}

$teamId = $team.GroupId

# Import the CSV file
$users = Import-Csv -Path $csvPath

# Loop through each user in the CSV and add them to the team
foreach ($user in $users) {
    $email = $user.UserEmail
    $role = $user.Role

    if ($role -eq "Owner") {
        # Add user as an Owner
        Add-TeamUser -GroupId $teamId -User $email -Role Owner
        Write-Host "Adding user $email to Microsoft Team $TeamName as $role"
    } elseif ($role -eq "Member") {
        # Add user as a Member
        Add-TeamUser -GroupId $teamId -User $email -Role Member
        Write-Host "Adding user $email to Microsoft Team $TeamName as $role"
    } else {
        Write-Host "Invalid role for $email"
    }
}

# Disconnect from Microsoft Teams
Disconnect-MicrosoftTeams
