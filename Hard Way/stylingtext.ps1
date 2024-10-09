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

# Define the email parameters
$recipientEmail = "npatil@aligntech.com" 
$subject = "Hello World"

# Create the email body with styling
$bodyContent = @"
<html>
<head>
<style>
    .blue-text {
        color: blue;
        font-weight: bold;
    }
    .green-text {
        color: green;
        font-size: 16px;
    }
    .red-text {
        color: red;
        font-size: 14px;
        font-style: italic;
    }
</style>
</head>
<body>
    <p class="blue-text">Hello, World!</p>
    <p class="green-text">This is a test email with styled content.</p>
    <p class="red-text">Enjoy sending styled emails!</p>
</body>
</html>
"@

# Create the email message
$emailMessage = @{
    Message         = @{
        Subject      = $subject
        Body         = @{
            ContentType = "HTML"
            Content     = $bodyContent
        }
        ToRecipients = @(
            @{
                EmailAddress = @{
                    Address = $recipientEmail
                }
            }
        )
    }
    SaveToSentItems = "true"
}

# Send the email
Send-MgUserMail -UserId $recipientEmail -BodyParameter $emailMessage  

Write-Host "Email sent successfully to $recipientEmail."
