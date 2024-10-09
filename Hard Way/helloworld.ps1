#region credentials
# Credentials for your app registration that has permissions for Graph API mail.send
$TenantID = "your_tenant_id"
$MySecureString = ConvertTo-SecureString (Get-Secret credAADOrchaSecret -AsPlainText) -AsPlainText -Force
$MyConnectCred = New-Object System.Management.Automation.PSCredential((Get-Secret credAADOrchaAppID -AsPlainText), $MySecureString)
$MyEmailCred = ConvertTo-GraphCredential -ClientID (Get-Secret credAADOrchaAppID -AsPlainText) -ClientSecret (Get-Secret credAADOrchaSecret -AsPlainText) -DirectoryID $TenantID

# First, authenticate to Azure
$AZConnect = Connect-AzAccount -ServicePrincipal -Tenant $TenantID -Credential $MyConnectCred

# Obtain an Azure access token specifically for Microsoft Graph
$AzAccessToken = (Get-AzAccessToken -ResourceTypeName MSGraph -AsSecureString -WarningAction SilentlyContinue).Token  

# Establishes a connection to Microsoft Graph using the access token obtained earlier
Connect-MgGraph -AccessToken $AzAccessToken -NoWelcome
#endregion

# Create the email message
$emailMessage = @{
    Message         = @{
        Subject      = "Hello World"
        Body         = @{
            ContentType = "Text"
            Content     = "Hello, World! This is a test email."
        }
        ToRecipients = @(
            @{
                EmailAddress = @{
                    Address = "npatil@aligntech.com"
                }
            }
        )
    }
    SaveToSentItems = "true"
}

# Send the email
Send-MgUserMail -UserId "npatil@aligntech.com" -BodyParameter $emailMessage  

Write-Host "Email sent successfully to $recipientEmail."
