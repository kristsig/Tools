## Authenticate with the SharePoint Online site
# Required site specific parameters
$SPOSiteUrl = 'https://companyname.sharepoint.com/'
$SPOUserName = 'someoney@companyname.onmicrosoft.com'
$SPOPassword = 'mypassword'

# Alternative method for password handling
#$SPOPassword = Read-Host -Prompt "Enter your password: " -AsSecureString