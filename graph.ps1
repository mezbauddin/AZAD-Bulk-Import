# Import required modules
Install-Module -Name Microsoft.Graph.Authentication -Scope CurrentUser -Force -AllowClobber
Install-Module -Name Microsoft.Graph.Users -Scope CurrentUser -Force -AllowClobber

# Import modules after installation
Import-Module Microsoft.Graph.Authentication
Import-Module Microsoft.Graph.Users

# Define Graph API credentials
$appId = "<Your-AppId>"
$tenantId = "<Your-TenantId>"
$clientId = "<Your-ClientId>"
$clientSecret = "<Your-ClientSecret>"
$redirectUri = "https://login.microsoftonline.com/common/oauth2/nativeclient"

# Connect to Microsoft Graph
Connect-MgGraph -ClientID $clientId -TenantId $tenantId -ClientSecret $clientSecret -RedirectUri $redirectUri

# Get Azure AD tenant details
$tenant = Get-MgOrganization

if ($tenant) {
    $tenantId = $tenant.id
    Write-Output "Azure AD Tenant ID: $tenantId"

    # Get the directory of the script
    $scriptDirectory = Split-Path -Parent $MyInvocation.MyCommand.Path

    # Read CSV and process users
    $users = Import-Csv -Path (Join-Path -Path $scriptDirectory -ChildPath "users.csv")

    # Loop through users
    foreach ($user in $users) {
        $email = $user.'Email'
        $firstName = $user.'First name'
        $lastName = $user.'Last name'
        $companyName = $user.'Company Name'
        $city = $user.'City'
        $department = $user.'Department'
        $employeeType = $user.'Employee Type'

        # Check if email is provided and is in a valid format
        if (-not ([string]::IsNullOrWhiteSpace($email)) -and $email -match '\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b') {
            # Check if user already exists in Azure AD
            $existingUser = Get-MgUser -Filter "mail eq '$email'"
            if ($existingUser) {
                Write-Output "User with email $email already exists in Azure AD. Checking for updates..."

                # Check for updates in user properties
                $updateRequired = $false
                if ($existingUser.givenName -ne $firstName) {
                    $existingUser.givenName = $firstName
                    $updateRequired = $true
                }
                if ($existingUser.surname -ne $lastName) {
                    $existingUser.surname = $lastName
                    $updateRequired = $true
                }
                if ($existingUser.companyName -ne $companyName) {
                    $existingUser.companyName = $companyName
                    $updateRequired = $true
                }
                if ($existingUser.city -ne $city) {
                    $existingUser.city = $city
                    $updateRequired = $true
                }
                if ($existingUser.department -ne $department) {
                    $existingUser.department = $department
                    $updateRequired = $true
                }
                if ($existingUser.employeeType -ne $employeeType) {
                    $existingUser.employeeType = $employeeType
                    $updateRequired = $true
                }
                if ($updateRequired) {
                    # Update user properties
                    Update-MgUser -UserId $existingUser.id -User $existingUser
                    Write-Output "User properties updated for $email."
                } else {
                    Write-Output "No updates found for user $email."
                }

                # TODO: Add user to the specified groups
            } else {
                Write-Output "User doesn't exist in Azure AD. Skipping user update."
            }
        } else {
            Write-Output "Invalid email address provided for user $firstName $lastName."
        }
    }
} else {
    Write-Output "Failed to retrieve Azure AD tenant details."
}
