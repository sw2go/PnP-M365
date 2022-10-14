# pnp-m365

sample code and setup for .net client to access a sharepoint site online (microsoft 365)

topics covered:
- client auth with selfsigned certificate
- limit sharepoint access with granular permission to one site only
- .net c# client (using PnP.Core)


First, use Powershell to create a selfsigned certificate for client authentication (*.cer for server and *.pfx for client)
https://github.com/sw2go/certificates-cs/blob/main/create-certificates.ps

Install the *.pfx on your development machine where you are going to use powershell and the .net client later on

Register a 1st Application in Azure
-----------------------------------
Open portal.azure.com, azure active directory, and register a application that will be used by the .net client
- New registration, name: sharepoint-test, check Accounts in any organizational directory (Any Azure AD directory - Multitenant)
- click register
- Authentication, add a platform, Mobile and desktop applications, Redirect URIs: https://login.microsoftonline.com/common/oauth2/nativeclient
- Certificates & secrets, upload the *.cer
- API-Permissions: Microsoft Graph, Application Permissions, Sites.Selected, User.Read (is default and can remain)
- API-Permissions: Sharepoint,      Application Permissions, Sites.Selected
- Grant admin consent

Register a 2nd Application in Azure (Admin-App to set the permissions for the 1st Application)
-----------------------------------
Open portal.azure.com, azure active directory, and register a application that will just be used by a Admin and powershell to setup granular sharepoint permission
- New registration, name: sharepoint-admin, check Accounts in any organizational directory (Any Azure AD directory - Multitenant)
- click register
- Authentication, add a platform, Desktop & Mobile, Redirect URIs: https://login.microsoftonline.com/common/oauth2/nativeclient
- Certificates & secrets, upload the *.cer
- API-Permissions: Microsoft Graph, Application Permissions, Sites.Fullcontrol.All, User.Read (is default and can remain)
- Grant admin consent

Run Powershell to administrate permissions the Client-Application has for the sharepoint site (the actual config behind Sites.Selected)
------------------------------------------
$tenantName = "contoso"

# First get a connection to the Admin-App where you gave Microsoft Graph, Application Permissions, Sites.FullControl.All
# Set the ClientId to the Azure ClientId of the Admin-App you registered before and the Thumbprint to the *.pfx you installed locally
$conn = Connect-PnPOnline -ClientId "51f5c579-2836-43bb-9af2-23939a296661" -Thumbprint 4D743BDB622CACF7653D7E97FB8C01639CE4F6EE -Tenant "$($tenantName).onmicrosoft.com" -Url "https://$($tenantName).sharepoint.com" -ReturnConnection

# Then use this connection to set the sharepoint site permissions for the 1st azure application
$spoSiteUrl = "https://$($tenantName).sharepoint.com/sites/1008"
$azureAppId = "1a97fc28-dc30-4a81-9a22-91b3af58ac9a"

# List die Permission-Commands
gcm -name "*PnPAzureADAppSite*"

# Read the current permissions
$perm = Get-PnPAzureADAppSitePermission -Site $spoSiteUrl -AppIdentity $azureAppId -Connection $conn

# Show the current permissions
$perm

# Grant Read Permission
Grant-PnPAzureADAppSitePermission       -Site $spoSiteUrl -AppId $azureAppId -DisplayName "theAppName" -Permissions Read -Verbose -Connection $conn

# Set Write Permission
Set-PnPAzureADAppSitePermission         -Site $spoSiteUrl -PermissionId $perm.Id -Permissions Write -Verbose -Connection $conn

#Revoke
Revoke-PnPAzureADAppSitePermission      -Site $spoSiteUrl -PermissionId $perm.Id -Force -Connection $conn

Create a .NET Client to access the Client-Application
-----------------------------------------------------
create new c# console project
Install NuGet PnP.Core.Auth, Microsoft.Extensions.Hosting
add appsettings.json (set copy to output always)
etc.
