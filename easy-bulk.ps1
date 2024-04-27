# Install AzureAD & Microsoft.Graph modules
Install-Module -Name "AzureAD" -Force -AllowClobber -SkipPublisherCheck
Install-Module -Name "Microsoft.Graph" -Force -AllowClobber -SkipPublisherCheck

# Import AzureAD module & Microsoft.Graph module
Import-Module -Name "AzureAD" -Force
Import-Module -Name "Microsoft.Graph" -Force

# Connect to AzureAD & Microsoft.Graph
Connect-AzureAD
Connect-MgGraph -Scopes Directory.ReadWrite.All

# Get Azure AD tenant details
$tenantDetail = Get-AzureADTenantDetail

if ($tenantDetail) {
    $tenantId = $tenantDetail.ObjectId
    Write-Output "Azure AD Tenant ID: $tenantId"

    # Define batch size and wait time (in seconds)
    $batchSize = 100
    $waitTime = 60  # Adjust as needed based on Azure AD performance and rate limits

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
            $existingUser = Get-AzureADUser -Filter "Mail eq '$email'"
            if ($existingUser) {
                Write-Output "User with email $email already exists in Azure AD. Checking for updates..."

                # Check for updates in user properties
                $updateRequired = $false
                if ($existingUser.GivenName -ne $firstName) {
                    $existingUser.GivenName = $firstName
                    $updateRequired = $true
                }
                if ($existingUser.Surname -ne $lastName) {
                    $existingUser.Surname = $lastName
                    $updateRequired = $true
                }
                if ($existingUser.CompanyName -ne $companyName) {
                    $existingUser.CompanyName = $companyName
                    $updateRequired = $true
                }
                if ($existingUser.City -ne $city) {
                    $existingUser.City = $city
                    $updateRequired = $true
                }
                if ($existingUser.Department -ne $department) {
                    $existingUser.Department = $department
                    $updateRequired = $true
                }
                if ($updateRequired) {
                    # Update user properties
                    Set-AzureADUser -ObjectId $existingUser.ObjectId -GivenName $existingUser.GivenName -Surname $existingUser.Surname -CompanyName $existingUser.CompanyName -City $existingUser.City -Department $existingUser.Department
                    
                    Write-Output "User properties updated for $email."
                } else {
                    Write-Output "No updates found for user $email."
                }
            } else {
                # User doesn't exist, send invitation
                $invitation = New-AzureADMSInvitation -InvitedUserEmailAddress $email -InviteRedirectUrl "https://myapps.microsoft.com" -SendInvitationMessage $false -InvitedUserDisplayName "$firstName $lastName" -InvitedUserType Guest

                if ($invitation) {
                    Write-Output "Invitation sent to $email for $firstName $lastName."

                    # Wait for a moment before proceeding
                    # Start-Sleep -Seconds 10  # Adjust if necessary

                    # Display a message to prompt to check email for invitation

                    # Proceed to the next user
                    continue
                } else {
                    Write-Output "Failed to send invitation to $email for $firstName $lastName."
                    # Proceed to the next user
                    continue
                }
            }

            # Wait for a moment before proceeding to ensure user creation completes
            # Start-Sleep -Seconds 10  # Adjust if necessary
            
            # Get the newly created user object
            $newUser = Get-AzureADUser -Filter "Mail eq '$email'"

            if ($newUser) {
                # Update user properties
                $userParams = @{
                    GivenName = $firstName
                    Surname = $lastName
                    CompanyName = $companyName
                    City = $city
                    Department = $department
                    # EmployeeType = $employeeType
                }
                Set-AzureADUser -ObjectId $newUser.ObjectId @userParams

                # Update employee type using Microsoft Graph
                Update-MgUser -UserId $newUser.UserPrincipalName -EmployeeType $employeeType
                Write-Output "Employee type updated for $email."

                # Add user to the specified groups
                for ($j = 1; $j -le 2; $j++) {
                    $groupName = $user."Group $j"
                    if (-not [string]::IsNullOrWhiteSpace($groupName)) {
                        $group = Get-AzureADGroup -Filter "DisplayName eq '$groupName'"
                        if ($group) {
                            try {
                                Add-AzureADGroupMember -ObjectId $group.ObjectId -RefObjectId $newUser.ObjectId -ErrorAction Stop
                                Write-Output "User added to group $groupName."
                            } catch {
                                if ($_.Exception.Message -match 'members') {
                                    Write-Output "User is already a member of group $groupName. Skipping..."
                                } else {
                                    Write-Output "Failed to add user to group $($groupName): $($_.Exception.Message)"
                                }
                            }
                        } else {
                            Write-Output "Failed to find group $groupName for user with email $email."
                        }
                    }
                }
            } else {
                Write-Output "Failed to retrieve newly created user object for $email."
            }
        } else {
            Write-Output "Invalid email address provided for user $firstName $lastName."
        }
    }
} else {
    Write-Output "Failed to retrieve Azure AD tenant details."
}

# Disconnect from Azure AD & Microsoft Graph
Disconnect-AzureAD
Disconnect-MgGraph
