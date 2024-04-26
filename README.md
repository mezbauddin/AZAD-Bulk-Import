# Bulk Import Users to Azure Active Directory

This PowerShell script enables bulk import of users into Azure Active Directory (Azure AD) and updates their properties using a CSV file. It utilizes the AzureAD and Microsoft.Graph PowerShell modules to interact with Azure AD and Microsoft Graph API.

## Prerequisites

Before running the script, ensure you have the following prerequisites:

- PowerShell installed on your machine.
- AzureAD and Microsoft.Graph PowerShell modules installed. You can install them using the following commands:

```powershell
# Install AzureAD module
Install-Module -Name AzureAD -Scope CurrentUser -Force

# Install Microsoft.Graph module
Install-Module -Name Microsoft.Graph -Scope CurrentUser -Force
Permissions to manage users and groups in your Azure AD tenant.
Usage
Clone this repository to your local machine.
Place your users.csv file containing user data in the same directory as the script.
Modify the users.csv file with the user data you want to import. Make sure the column names match the properties you want to update.
Open PowerShell and navigate to the directory where the script is located.
Run the script using the following command:
powershell
Copy code
.\bulk.ps1
CSV File Format
The CSV file should contain the following columns:

Email
First name
Last name
Company Name
City
Department
Employee Type
Group 1
Group 2
Note: You can customize the column names and adjust the script accordingly.

Contributing
Contributions are welcome! If you find any issues or have suggestions for improvements, feel free to open an issue or submit a pull request.

License
This project is licensed under the MIT License.
