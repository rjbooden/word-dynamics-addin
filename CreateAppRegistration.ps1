# Create App registration

$appName = "Word Dynamics Add-In"
$readName = Read-Host -Prompt "Please supply the application name (default: $appName)"
if (-not [string]::IsNullOrWhiteSpace($readName)) {
    $appName = $readName
}

$hostName = "localhost:7150"
$validUrl = $false
while (-not $validUrl) {
    $readUrl = Read-Host -Prompt "Please supply the url where the application will be hosted (include port if not 443)"
    $validUrl = $readUrl.StartsWith("https://")
    if ($validUrl) {
        $hostName = $readUrl -replace "https://", ""
        $hostName = $hostName.Trim("/")
    }
    else {
        Write-Host "Use https://yoururl"
    }
}

# See: https://github.com/microsoftgraph/microsoft-graph-docs/blob/main/concepts/includes/permissions-ids.md
$resourceAccess = (@(@{ 
    resourceAppId = "00000003-0000-0000-c000-000000000000" # Graph API
    resourceAccess = @(@{
        id = "14dad69e-099b-42c9-810b-d002981feec1" # profile
        type = "Scope"
    },
    @{
        id = "e1fe6dd8-ba31-4d61-89e7-88639da4683d" # User.Read
        type = "Scope"
    },
    @{
        id = "5c28f0bf-8a70-41f1-8ab2-9032436ddb65" # Files.ReadWrite
        type = "Scope"
    },
    @{
        id = "37f7f235-527c-4136-accd-4a02d197296e" # openid
        type = "Scope"
    })
    
},@{
    resourceAppId = "00000007-0000-0000-c000-000000000000" # Dynamics CRM
    resourceAccess = @(
        @{
            id = "78ce3f0f-a1ce-49c2-8cde-64b5c0896db4" # user_impersonation
            type = "Scope"
        }
    )
}) | ConvertTo-Json -Depth 4 -Compress) -replace "`"","\`"" 

$permissionGuid = [guid]::NewGuid()
$additionalSettings = (@{
    api = @{
        requestedAccessTokenVersion = 2
        oauth2PermissionScopes = @(@{
            adminConsentDescription = "access_as_user"
            adminConsentDisplayName = "access_as_user"
            id = $permissionGuid
            isEnabled = $true
            #origin = "Application"
            type = "Admin"
            value = "access_as_user"
        })
    }
} | ConvertTo-Json -Depth 4 -Compress) -replace "`"","\`"" 

$additionalSettings2 = (@{
    api = @{
        preAuthorizedApplications = @(
        @{
            appId = "ea5a67f6-b6f3-4338-b240-c655ddc3cc8e" # SSO Office Add-Ins
            delegatedPermissionIds = @($permissionGuid)
        })
    }
} | ConvertTo-Json -Depth 4 -Compress) -replace "`"","\`"" 

$webRedirectUri = "https://$hostName/signin-oidc"
Write-Host "Adding app registration for: $appName"
$app = az ad app create --display-name $appName --sign-in-audience AzureADMyOrg --web-redirect-uris $webRedirectUri --required-resource-accesses $resourceAccess | ConvertFrom-Json
$clientId = $app.appId
$objectId = $app.id
$identifierUri = "api://$hostName/$clientId"
az ad app update --id $objectId  --identifier-uris $identifierUri | Out-Null
$uri = "https://graph.microsoft.com/v1.0/applications/$objectId"

Write-Host "Adding permissions and scope"
az rest --method PATCH --uri $uri --headers "`"Content-Type`"=`"application/json`"" --body $additionalSettings
Write-Host "Pre authorize Office for using SSO in Office Add-Ins"
# Needs second step as oauth2PermissionScopes needs to be set first
az rest --method PATCH --uri $uri --headers "`"Content-Type`"=`"application/json`"" --body $additionalSettings2

Write-Host "Please wait for the app registration to be processed."
# Wait
Start-Sleep 30

$readKey = Read-Host -Prompt "Apply Admin consent now (y/N)?"
if ($readKey -ieq "y") {
    az ad app permission admin-consent --id $objectId
}
else {
    Write-Host "Please give admin consent for the 'API permissions' of the '$appName' App registraion in the Azure portal"
}

Write-Host "Creating an retieving app secret"
# Add and get secret to app registration
$secretResponse = az rest --method post --uri "https://graph.microsoft.com/v1.0/applications/$objectId/addPassword" | ConvertFrom-Json
Write-Host "Your app secret is (store securely): $($secretResponse.secretText)"
Write-Host "Please remember it expires on: $($secretResponse.endDateTime)"
Write-Host "Your client id is: $clientId"
Write-Host "Identyfier uri: $identifierUri"

# Update config settings
$readKey = Read-Host -Prompt "Would you like to configure your appsettings.json and manifest.xml (Y/n)?"
if ($readKey -ine "n") {
    $validPath = $false
    $validDynamicsName = $false
    while (-not $validPath) {
        $readLine = Read-Host "Please provide the path the published root containing the appsettings.json file"
        if([string]::IsNullOrWhiteSpace($readLine)) {
            Exit
        }
        $validPath = (Test-Path $readLine\appsettings.json) -and (Test-Path $readLine\wwwroot\manifest.xml)
        if (-not $validPath) {
            Write-Host "Could not locate the appsettings.json and manifest.xml given the path: $readLine"
        }
    }

    while (-not $validDynamicsName) {
        $readDynamicsName = Read-Host "What is your Dynamics name (https://[Dynamics name].crm4.dynamics.com)"
        $validDynamicsName = [string]::IsNullOrWhiteSpace($readDynamicsName) -ne $true
        if ($readDynamicsName -ieq "exit") {
            Exit
        }
        if (-not $validDynamicsName) {
            Write-Host "Please provide the Dynamics name or type exit to stop this configuration"
        }
    }

    $tenantInfo = az account show | ConvertFrom-Json
    $tenantId = $tenantInfo.homeTenantId
    $domains = az rest --method get --url 'https://graph.microsoft.com/v1.0/domains' | ConvertFrom-Json
    $defaultTenant = $domains.value | Where-Object { $_.isDefault }
    $tenantDomain = $defaultTenant.id

    $appSettings = Get-Content -Path $readLine\appsettings.json
    $appSettings = $appSettings -replace "\[your domain]", "$tenantDomain"
    $appSettings = $appSettings -replace "\[your TenantId\]", "$tenantId"
    $appSettings = $appSettings -replace "\[your ClientId\]", "$clientId"
    $appSettings = $appSettings -replace "\[your secret here\]", "$($secretResponse.secretText)"
    $appSettings = $appSettings -replace "\[your_dynamics\]", "$readDynamicsName"
    Set-Content -Path $readLine\appsettings.json -Value $appSettings -Force

    $manifest = Get-Content -Path $readLine\wwwroot\manifest.xml
    $manifest = $manifest -replace "\[your_url\]", "$hostName"
    $manifest = $manifest -replace "\[your clientid\]", "$clientId"
    Set-Content -Path $readLine\wwwroot\manifest.xml -Value $manifest -Force

    Write-Host "Configuration done..."
}