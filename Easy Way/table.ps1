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

#defining a table 
$table = @([PSCustomObject]@{
        item        = "Item1"
        Description = "Description1"
        price       = "10.00"
    }, [PSCustomObject]@{
        item        = "Item2"
        Description = "Description2"
        price       = "20.00"
    }, [PSCustomObject]@{
        item        = "Item3"
        Description = "Description3"
        price       = "30.00"
    })

$EmailBody = EmailBody {
    EmailText -Text "Hello, World!" -LineBreak -Color Purple -FontWeight Bold
    EmailText -Text "This is a test email with styled content." -LineBreak -Color Green -FontSize 16
    EmailText -Text "Enjoy sending styled emails!" -LineBreak -Color Red -FontSize 14 -FontStyle italic
    EmailTable -Table $table -HideFooter 
}
$MailParams = @{
    To      = "npatil@aligntech.com"
    From    = "npatil@aligntech.com"
    Subject = "Hello World"
    Body    = $EmailBody
}
Send-EmailMessage @MailParams -Credential $MyEmailCred -Graph 