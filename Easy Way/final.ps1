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

#region emailbody
$Body = EmailBody {
    #region text styling
    EmailText -Text "Hello Dear Reader," -LineBreak 
    EmailText -Text @(
        "Write your ", "text", " in ", "multiple ways: ", " colors", " or ", "fonts", " or ", "text transformations!"
    ) -Color Blue, Red, Yellow, GoldenBrown, SeaGreen, None, Green, None -FontWeight normal, bold, normal, bold, normal, normal, normal, normal, normal -LineBreak
    #endregion
    #region nested lists
    EmailText -Text "You can create lists: "
    EmailList {
        EmailListItem -Text "First item"
        EmailListItem -Text "Second item"
        EmailListItem -Text "Third item"
        EmailList {
            EmailListItem -Text "Nested item 1"
            EmailListItem -Text "Nested item 2"
        } -Type Unordered
    } -Type Ordered -FontSize 15
    #endregion
    #region tables
    EmailText -Text "You can create tables: " -LineBreak
    EmailTable -DataTable (Get-Process | Select-Object -First 5 -Property Name, Id, PriorityClass, CPU, Product) -HideFooter
    EmailText -LineBreak
    #endregion
    EmailText -Text "Everything is customizable. " -Color California -FontStyle italic -TextDecoration underline
    #region images
    EmailText -Text "You can even add images: " -LineBreak 
    EmailImage -Source "https://www.google.com/images/branding/googlelogo/2x/googlelogo_color_272x92dp.png"  -width 200 -height 100 
    #endregion
    EmailText -Text "It's all just a command away. " -Color None -FontStyle normal -TextDecoration none
    EmailText -Text "You no longer have to use HTML/CSS, as it will be used for you!"
}
#endregion

$MailParams = @{
    To          = "npatil@aligntech.com"
    From        = "npatil@aligntech.com"
    Subject     = "Hello World"
    Body        = $Body
    Attachments = "C:\Users\npatil\Downloads\Log_20240728_003353.csv"
    Priority    = "High"
}
Send-EmailMessage @MailParams -Credential $MyEmailCred -Graph -