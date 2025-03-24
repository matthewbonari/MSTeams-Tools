# Install required modules if not already installed
# Install-Module Microsoft.Graph -Scope CurrentUser
# Install-Module MicrosoftTeams -Scope CurrentUser

# Import required modules, commented out, already installed
# Import-Module Microsoft.Graph
# Import-Module MicrosoftTeams

# Connect to Microsoft Graph and Teams
Connect-MgGraph -Scopes "User.Read.All", "Directory.Read.All"
Connect-MicrosoftTeams

# Define input and output CSV file paths
$inputCsv = "C:\Temp\Input.csv"
$outputCsv = "C:\Temp\Output.csv"

# Read CSV file
$users = Import-Csv -Path $inputCsv

# Initialize output array
$outputData = @()

foreach ($user in $users) {
    #$userId = $user.Identity
    #try {
    #    $userID = Get-CsOnlineUser -Identity $user -ErrorAction Stop
    #} catch {
    #    $userID = "Not Found"
    #}

    # Retrieve user details
    $userDetails = Get-CsOnlineUser -Identity $user.ID

    
    # Retrieve managerId
    try {
        $manager = Get-MgUserManager -UserId $userDetails.UserPrincipalName -ErrorAction Stop
        $managerId = $manager.Id
    } catch {
        $managerId = "Not Found"
    }
    
    # Retrieve manager details
    try {
        $managerDetails = Get-CsOnlineUser -Identity $managerId -ErrorAction Stop
        $managerUPN = $managerDetails.UserPrincipalName
    } catch {
        $managerUPN = "Not Found"
    }
    
    # Store data
    $outputData += [PSCustomObject]@{
        Identity           = $userId
        UserPrincipalName  = $userDetails.UserPrincipalName
        ManagerId          = $managerId
        ManagerUPN         = $managerUPN
    }
}

# Export results to CSV
$outputData | Export-Csv -Path $outputCsv -NoTypeInformation

Write-Host "Process completed. Output saved to: $outputCsv"
