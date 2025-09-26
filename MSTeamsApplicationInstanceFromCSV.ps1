<#
.VERSION
1.0

.AUTHOR 
Matthew Bonari

.SYNOPSIS
Creates Microsoft Teams Application Instances from CSV data for building Call Queues and Auto Attendants
Update $CsvPath variable to match .CSV file location
Authenticate with Microsoft Teams module prior to running script using Connect-MicrosoftTeams

.DESCRIPTION
This script contains common commands for reference

#.PARAMETER -UserPrincipalName
Email at domain

#.EXAMPLE
New-CsOnlineApplicationInstance -UserPrincipalName "lobbyphone@contoso.com" -DisplayName "Lobby Phone" -ApplicationId "ce933385-9390-45d1-9512-c8d228074e07"

#.NOTES
Learn documentation for new-csonlineapplicationinstance - https://learn.microsoft.com/en-us/powershell/module/microsoftteams/new-csonlineapplicationinstance?view=teams-ps

#>

# Global Variables
$AaApplicationId = "ce933385-9390-45d1-9512-c8d228074e07"
$CqApplicationId = "11cd3e2e-fccb-42ad-ad00-878b93575e07"
$CsvPath = "C:\temp\filename.csv"

# Check Module Dependancy
try {
    Write-Host "Checking for MicrosoftTeams module..." -ForegroundColor Yellow
    if (-not (Get-Module -Name MicrosoftTeams -ListAvailable)) {
        throw "MicrosoftTeams module not found. Please install with: Install-Module MicrosoftTeams"
    }
    Import-Module MicrosoftTeams -Force
    Write-Host "✓ MicrosoftTeams module loaded successfully" -ForegroundColor Green
}
catch {
    Write-Error "Failed to load MicrosoftTeams module: $($_.Exception.Message)"
    exit 1
}


# Import and validate CSV data
try {
    Write-Host "Importing CSV data from: $CsvPath" -ForegroundColor Yellow
    $CsvData = Import-Csv -Path $CsvPath -ErrorAction Stop
    
    if ($CsvData.Count -eq 0) {
        throw "CSV file is empty or contains no valid data rows"
    }
    
    # Validate required columns
    $RequiredColumns = @('UserPrincipalName', 'DisplayName', 'ApplicationId')
    $CsvColumns = $CsvData[0].PSObject.Properties.Name
    
    foreach ($Column in $RequiredColumns) {
        if ($Column -notin $CsvColumns) {
            throw "Required column '$Column' not found in CSV. Available columns: $($CsvColumns -join ', ')"
        }
    }
    
    Write-Host "✓ CSV validation successful. Found $($CsvData.Count) records" -ForegroundColor Green
}
catch {
    Write-Error "CSV import failed: $($_.Exception.Message)"
    Disconnect-MicrosoftTeams -Confirm:$false
    exit 1
}

# Initialize tracking variables
$SuccessCount = 0
$FailureCount = 0
$Results = @()

# Process each row from CSV
Write-Host "`nProcessing application instances..." -ForegroundColor Yellow

foreach ($Row in $CsvData) {
    $CurrentUPN = $Row.UserPrincipalName
    
    try {
        # Validate row data
        if ([string]::IsNullOrWhiteSpace($Row.UserPrincipalName)) {
            throw "UserPrincipalName is empty or null"
        }
        if ([string]::IsNullOrWhiteSpace($Row.DisplayName)) {
            throw "DisplayName is empty or null"
        }
        if ([string]::IsNullOrWhiteSpace($Row.ApplicationId)) {
            throw "ApplicationId is empty or null"
        }
        
        Write-Host "Creating: $($Row.DisplayName) ($CurrentUPN)" -ForegroundColor Cyan
        
        # Create the application instance
        $AppInstance = New-CsOnlineApplicationInstance `
            -UserPrincipalName $Row.UserPrincipalName `
            -DisplayName $Row.DisplayName `
            -ApplicationId $Row.ApplicationId `
            -ErrorAction Stop
        
        Write-Host "✓ Success: $($Row.DisplayName)" -ForegroundColor Green
        $SuccessCount++
        
        # Store result
        $Results += [PSCustomObject]@{
            UserPrincipalName = $Row.UserPrincipalName
            DisplayName = $Row.DisplayName
            ApplicationId = $Row.ApplicationId
            Status = "Success"
            ObjectId = $AppInstance.ObjectId
            Error = $null
        }
    }
    catch {
        $ErrorMessage = $_.Exception.Message
        Write-Host "✗ Failed: $CurrentUPN - $ErrorMessage" -ForegroundColor Red
        $FailureCount++
        
        # Store error result
        $Results += [PSCustomObject]@{
            UserPrincipalName = $CurrentUPN
            DisplayName = $Row.DisplayName
            ApplicationId = $Row.ApplicationId
            Status = "Failed"
            ObjectId = $null
            Error = $ErrorMessage
        }
    }
}