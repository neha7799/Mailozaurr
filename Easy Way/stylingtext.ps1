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

#scriptblock to define the email body
$EmailBody = EmailBody {
    EmailText -Text "Hello, World!" -LineBreak -Color Purple -fontweight bold 
    EmailText -Text "This is a test email with styled content." -LineBreak -Color Green -FontSize 16
    EmailText -Text "Enjoy sending styled emails!" -LineBreak -Color Red -FontSize 14 -FontStyle italic
}
#splatting the parameters
$MailParams = @{
    To      = "npatil@aligntech.com"
    From    = "npatil@aligntech.com"
    Subject = "Hello World"
    Body    = $EmailBody
}
Send-EmailMessage @MailParams -Credential $MyEmailCred -Graph 