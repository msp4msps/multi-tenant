###Finds All Licensed Shared Mailboxes and Remvoes the License##



$ApplicationId = "Enter ApplicationId"
$ApplicationSecret = "Enter ApplicationSecret"
$secPas = $ApplicationSecret| ConvertTo-SecureString -AsPlainText -Force
$tenantID = "Enter TenantId"
$refreshToken = 'Enter refreshToken'
$upn = "Enter UPN"
$ExchangeRefreshToken = '"Enter ExchangeRefreshToken"'
   
$credential = New-Object System.Management.Automation.PSCredential($ApplicationId, $secPas)
 
$aadGraphToken = New-PartnerAccessToken -ApplicationId $ApplicationId -Credential $credential -RefreshToken $refreshToken -Scopes 'https://graph.windows.net/.default' -ServicePrincipal -Tenant $tenantID
$graphToken = New-PartnerAccessToken -ApplicationId $ApplicationId -Credential $credential -RefreshToken $refreshToken -Scopes 'https://graph.microsoft.com/.default' -ServicePrincipal -Tenant $tenantID
 
Connect-MsolService -AdGraphAccessToken $aadGraphToken.AccessToken -MsGraphAccessToken $graphToken.AccessToken
 
$customers = Get-MsolPartnerContract -All
 
Write-Host "Found $($customers.Count) customers for $((Get-MsolCompanyInformation).displayname)." -ForegroundColor DarkGreen
 
foreach ($customer in $customers) {

     #Get ALL Licensed Users and Find Shared  Mailboxes#
    try{
    $licensedUsers = Get-MsolUser -TenantId $customer.TenantId | Where-Object {$_.islicensed}
    Write-Host "Checking Shared Mailboxes for $($Customer.Name)" -ForegroundColor Green
    $token = New-PartnerAccessToken -ApplicationId 'a0c73c16-a7e3-4564-9a95-2bdf47383716'-RefreshToken $ExchangeRefreshToken -Scopes 'https://outlook.office365.com/.default' -Tenant $customer.TenantId -ErrorAction SilentlyContinue
    $tokenValue = ConvertTo-SecureString "Bearer $($token.AccessToken)" -AsPlainText -Force
    $credential = New-Object System.Management.Automation.PSCredential($upn, $tokenValue)
    $customerId = $customer.DefaultDomainName
    $InitialDomain = Get-MsolDomain -TenantId $customer.TenantId | Where-Object {$_.IsInitial -eq $true}
    $session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://ps.outlook.com/powershell-liveid?DelegatedOrg=$($InitialDomain)&BasicAuthToOAuthConversion=true" -Credential $credential -Authentication Basic -AllowRedirection -ErrorAction SilentlyContinue
    Import-PSSession $session 
    $sharedMailboxes = Get-Mailbox -ResultSize Unlimited -Filter {recipienttypedetails -eq "SharedMailbox"} | Get-MailboxStatistics | Where-Object {[int64]($PSItem.TotalItemSize.Value -replace '.+\(|bytes\)') -lt "50GB"}
    Remove-PSSession $session
    foreach ($mailbox in $sharedMailboxes) {
        if ($licensedUsers.displayName -contains $mailbox.displayName) {
            Write-Host "$($mailbox.displayname) is a licensed shared mailbox" -ForegroundColor Yellow
            $user = ($licensedUsers | Where-Object {$_.displayName -contains $mailbox.displayName})
            $licenses = ($licensedUsers | Where-Object {$_.displayName -contains $mailbox.displayName}).Licenses
            $licenseArray = $licenses | foreach-Object {$_.AccountSkuId}
            Write-Host "Removing License" -ForegroundColor Yellow
            $mailbox | ForEach-Object {
            Set-MsolUserLicense -UserPrincipalName "$($user.UserPrincipalName)" -TenantId $($customer.TenantId) -removelicenses $licenseArray -ErrorAction SilentlyContinue
            } 

}
}
}catch { "An error occurred."}
}
