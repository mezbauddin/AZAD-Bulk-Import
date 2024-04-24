# Check if AzureAD module is installed, if not install it
if (-not (Get-Module -Name AzureAD -ErrorAction SilentlyContinue)) {
    Write-Output "AzureAD module is not installed. Installing..."
    Install-Module -Name AzureAD -Scope CurrentUser -Force
}

# Import AzureAD module
Import-Module -Name AzureAD -Force

# Connect to Azure AD
Connect-AzureAD

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

    # Read CSV and create users
    $users = Import-Csv -Path (Join-Path -Path $scriptDirectory -ChildPath "users.csv")

    # Calculate number of batches
    $totalUsers = $users.Count
    $totalBatches = [math]::Ceiling($totalUsers / $batchSize)

    # Loop through batches
    for ($i = 0; $i -lt $totalBatches; $i++) {
        # Get current batch range
        $startIdx = $i * $batchSize
        $endIdx = [math]::Min(($startIdx + $batchSize - 1), ($totalUsers - 1))
        
        # Process current batch of users
        $batchUsers = $users[$startIdx..$endIdx]
        foreach ($user in $batchUsers) {
            $email = $user.'Current Work Email'
            $firstName = $user.'First name'
            $lastName = $user.'Last name'

            if (-not ([string]::IsNullOrWhiteSpace($email)) -and $email -match '\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b') {
                $invitation = New-AzureADMSInvitation -InvitedUserEmailAddress $email -InviteRedirectUrl "https://myapps.microsoft.com" -SendInvitationMessage $false -InvitedUserDisplayName "$firstName $lastName" -InvitedUserType Guest

                if ($invitation) {
                    Write-Output "Invitation sent to $email for $firstName $lastName."
                    $userObjectId = $invitation.InvitedUser.Id

                    # Update user properties
                    $userParams = @{
                        GivenName = $firstName
                        Surname = $lastName
                    }

                    Set-AzureADUser -ObjectId $userObjectId @userParams

                    # Add user to the specified groups
                    for ($j = 1; $j -le 2; $j++) {
                        $groupName = $user."Group $j"
                        if (-not [string]::IsNullOrWhiteSpace($groupName)) {
                            $group = Get-AzureADGroup -Filter "DisplayName eq '$groupName'"
                            if ($group) {
                                Add-AzureADGroupMember -ObjectId $group.ObjectId -RefObjectId $userObjectId
                                Write-Output "User added to group $groupName."
                            } else {
                                Write-Output "Failed to find group $groupName for user $firstName $lastName."
                            }
                        }
                    }
                } else {
                    Write-Output "Failed to send invitation to $email for $firstName $lastName."
                }
            } else {
                Write-Output "Invalid email address provided for user $firstName $lastName."
            }
        }
        
        # Output progress message
        Write-Output "Batch $($i + 1) of $totalBatches processed."
        
        # Introduce wait time between batches
        if ($i -lt ($totalBatches - 1)) {
            Write-Output "Waiting for $waitTime seconds before processing next batch..."
            Start-Sleep -Seconds $waitTime
        }
    }
} else {
    Write-Output "Failed to retrieve Azure AD tenant details."
}
