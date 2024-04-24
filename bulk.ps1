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

if ($tenantDetail -ne $null) {
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
            $password = ConvertTo-SecureString -String $user.Password -AsPlainText -Force

            # Check if user already exists
            $existingUser = Get-AzureADUser -Filter "UserPrincipalName eq '$($user.'Current Work Email')'"
            if ($existingUser -eq $null) {
                # Invite user as an external guest
                $displayName = "$($user.'First Name') $($user.'Last Name') - $($user.'Position Code') at $($user.'Co Name') in $($user.'Co City')"
                $invitation = New-AzureADMSInvitation -InvitedUserEmailAddress $user.'Current Work Email' -InviteRedirectUrl "https://myapps.microsoft.com" -SendInvitationMessage $false -InvitedUserDisplayName $displayName -InvitedUserType Guest

                if ($invitation -ne $null) {
                    Write-Output "Invitation sent to $($user.'Current Work Email') for $($user.'First Name') $($user.'Last Name')."
                }
                else {
                    Write-Output "Failed to send invitation to $($user.'Current Work Email') for $($user.'First Name') $($user.'Last Name')."
                }
            }
            else {
                Write-Output "User $($user.'Current Work Email') already exists in Azure AD. Skipping invitation."
            }

            # Add user to security groups
            for ($j = 1; $j -le 2; $j++) {
                $groupName = $user."Group $j"
                if (-not [string]::IsNullOrWhiteSpace($groupName)) {
                    $group = Get-AzureADGroup -Filter "DisplayName eq '$groupName'"
                    if ($group -ne $null) {
                        Add-AzureADGroupMember -ObjectId $group.ObjectId -RefObjectId $invitation.InvitedUser.ObjectId
                        Write-Output "User $($user.'First Name') $($user.'Last Name') added to group $groupName."
                    }
                    else {
                        Write-Output "Failed to find group $groupName for user $($user.'First Name') $($user.'Last Name')."
                    }
                }
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
}
else {
    Write-Output "Failed to retrieve Azure AD tenant details."
}
