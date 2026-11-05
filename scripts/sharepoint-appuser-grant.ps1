param(
    [Parameter(Mandatory)] [string] $TenantId,
    [Parameter(Mandatory)] [string] $AppName,
    [Parameter(Mandatory)] [string] $AppId,
    [Parameter(Mandatory)] [string] $SiteUrl
)

$ErrorActionPreference = 'Stop'

# This administrative session grants site access to the target app.
Connect-MgGraph -TenantId $TenantId `
    -Scopes 'Sites.FullControl.All' -ContextScope Process -NoWelcome

try {
    $siteUri = [uri]$SiteUrl
    $lookup = 'https://graph.microsoft.com/v1.0/sites/{0}:{1}' -f `
        $siteUri.Host, $siteUri.AbsolutePath.TrimEnd('/')

    $site = Invoke-MgGraphRequest -Method GET -Uri $lookup
    $permissionsUrl = "https://graph.microsoft.com/v1.0/sites/$($site.id)/permissions"

    # Find existing grants, including paginated results.
    $permissions = @()
    $next = $permissionsUrl
    while ($next) {
        $page = Invoke-MgGraphRequest -Method GET -Uri $next
        $permissions += @($page.value)
        $next = $page.'@odata.nextLink'
    }

    $existing = @($permissions | Where-Object {
        $identities = @($_.grantedToIdentitiesV2) +
                      @($_.grantedToIdentities)
        @($identities | Where-Object {
            $_.application.id -eq $AppId
        }).Count -gt 0
    })

    if ($existing.Count -gt 1) {
        throw "Multiple grants found for $AppName; review before changing."
    }

    if ($existing.Count -eq 1) {
        $grant = $existing[0]
        if ($grant.roles -notcontains 'fullcontrol') {
            $body = @{ roles = @('fullcontrol') } | ConvertTo-Json
            Invoke-MgGraphRequest -Method PATCH `
                -Uri "$permissionsUrl/$($grant.id)" `
                -Body $body -ContentType 'application/json' | Out-Null
        }
    }
    else {
        $body = @{
            roles = @('fullcontrol')
            grantedToIdentities = @(
                @{ application = @{
                    id          = $AppId
                    displayName = $AppName
                }}
            )
        } | ConvertTo-Json -Depth 6

        $grant = Invoke-MgGraphRequest -Method POST `
            -Uri $permissionsUrl -Body $body `
            -ContentType 'application/json'
    }

    $verified = Invoke-MgGraphRequest -Method GET `
        -Uri "$permissionsUrl/$($grant.id)"

    if ($verified.roles -notcontains 'fullcontrol') {
        throw "Full Control readback failed for $AppName."
    }

    [pscustomobject]@{
        Application = $AppName
        Site        = $site.webUrl
        Permission  = $verified.roles -join ', '
    }
}
finally {
    Disconnect-MgGraph | Out-Null
}
