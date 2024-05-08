# Function to check if a module is installed and import it if necessary
function Import-ModuleIfNeeded {
    param(
        [string]$ModuleName
    )
    
    if (-not (Get-Module -Name $ModuleName -ErrorAction SilentlyContinue)) {
        Write-Output "$ModuleName module is not installed. Installing..."
        Install-Module -Name $ModuleName -Scope CurrentUser -Force -AllowClobber
    } elseif (-not (Get-Module -Name $ModuleName)) {
        Import-Module -Name $ModuleName -Force
    } else {
        Write-Output "$ModuleName module is already imported."
    }
}

# Function to remove a user from Azure AD based on email address
function Remove-AzureADUser {
    param(
        [string]$EmailAddress
    )

    try {
        # Debugging output to verify email address
        Write-Output "Attempting to remove user with email: $EmailAddress"

        # Get the user by filtering on their email address
        $existingUser = Get-AzureADUser -Filter "Mail eq '$EmailAddress'"

        if ($existingUser) {
            # Remove the user
            Remove-AzureADUser -ObjectId $existingUser.ObjectId
            Write-Output "User with email $EmailAddress has been successfully removed from Azure AD."
        } else {
            Write-Output "User with email $EmailAddress does not exist in Azure AD."
        }
    } catch {
        Write-Output "Error: $_"
    }
}


# Check if AzureAD module is installed, if not install it
Import-ModuleIfNeeded -ModuleName "AzureAD"

# Connect to Azure AD
Connect-AzureAD

# Get the directory of the script
$scriptDirectory = Split-Path -Parent $MyInvocation.MyCommand.Path

# Read CSV and process users for removal
$usersToRemove = Import-Csv -Path (Join-Path -Path $scriptDirectory -ChildPath "users.csv")

# Loop through users to remove
foreach ($user in $usersToRemove) {
    $emailToRemove = $user.'Email'
    # Check if email is provided and is in a valid format
    if (-not ([string]::IsNullOrWhiteSpace($emailToRemove)) -and $emailToRemove -match '\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b') {
        # Remove user from Azure AD
        Remove-AzureADUser -EmailAddress $emailToRemove
    } else {
        Write-Output "Invalid email address provided for removal: $emailToRemove"
    }
}
