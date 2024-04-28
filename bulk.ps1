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

# Check if AzureAD module is installed, if not install it
Import-ModuleIfNeeded -ModuleName "AzureAD"

# Check if Microsoft.Graph.Users module is installed, if not install it
Import-ModuleIfNeeded -ModuleName "Microsoft.Graph.Users"

# Connect to Azure AD
Connect-AzureAD

# Connect to Microsoft Graph
Connect-MgGraph -Scopes Directory.ReadWrite.All


# Check if the connection to Microsoft Graph is successful
Write-Output "Successfully connected to Microsoft Graph."

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
                # Update employee type using Microsoft Graph
                Update-MgUser -UserId $existingUser.UserPrincipalName -EmployeeType $employeeType
                Write-Output "Employee type updated for $email."
                # Add user to the specified groups
                for ($j = 1; $j -le 2; $j++) {
                    $groupName = $user."Group $j"
                    if (-not [string]::IsNullOrWhiteSpace($groupName)) {
                        $group = Get-AzureADGroup -Filter "DisplayName eq '$groupName'"
                        if ($group) {
                            # Check if user is already a member of the group
                            $groupMembers = Get-AzureADGroupMember -ObjectId $group.ObjectId | Select-Object -ExpandProperty ObjectId
                            if ($groupMembers -notcontains $existingUser.ObjectId) {
                                try {
                                    Add-AzureADGroupMember -ObjectId $group.ObjectId -RefObjectId $existingUser.ObjectId -ErrorAction Stop
                                    Write-Output "User added to group $groupName."
                                } catch {
                                    if ($_.Exception.Message -match 'members') {
                                        Write-Output "User is already a member of group $groupName. Skipping..."
                                    } else {
                                        Write-Output "Failed to add user to group $($groupName): $($_.Exception.Message)"
                                    }
                                }
                            } else {
                                Write-Output "User is already a member of group $groupName. Skipping..."
                            }
                        } else {
                            Write-Output "Failed to find group $groupName for user with email $email."
                        }
                    }
                }
            } else {
                # User doesn't exist, send invitation
                $invitation = New-AzureADMSInvitation -InvitedUserEmailAddress $email -InviteRedirectUrl "https://myapps.microsoft.com" -SendInvitationMessage $false -InvitedUserDisplayName "$firstName $lastName" -InvitedUserType Guest

                if ($invitation) {
                    Write-Output "Invitation sent to $email for $firstName $lastName."

                    # Wait for a moment before proceeding
                    # Start-Sleep -Seconds 10  # Adjust if necessary
                   

                    # Get the newly created user object
                    $newUser = Get-AzureADUser -ObjectId $invitation.InvitedUser.Id

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
                    Write-Output "Failed to send invitation to $email for $firstName $lastName."
                }
            }
        } else {
            Write-Output "Invalid email address provided for user $firstName $lastName."
        }
    }
} else {
    Write-Output "Failed to retrieve Azure AD tenant details."
}

# Disconnect from MgGraph
# Disconnect-MgGraph