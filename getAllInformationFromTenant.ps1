# Configuration
$tenantId = ""
$clientId = ""
$clientSecret = ""

# Create Report Structure
$reportFolders = @{
    "Users"      = @(
        "BasicInfo",
        "GuestAccounts",
        "InactiveAccounts",
        "MFAStatus",
        "ProductUsage"
    )
    "Teams"      = @(
        "BasicInfo",
        "GuestAccess",
        "Channels",
        "Usage"
    )
    "SharePoint" = @(
        "Sites",
        "Storage",
        "Usage"
    )
    "Licenses"   = @(
        "Assigned",
        "History",
        "Usage",
        "ServicePlans"
    )
    "Security"   = @(
        "Score",
        "ConditionalAccess",
        "DeviceCompliance"
    )
}

# Create folder structure
foreach ($folder in $reportFolders.Keys) {
    $path = Join-Path "Reports" $folder
    New-Item -ItemType Directory -Force -Path $path | Out-Null
    foreach ($subFolder in $reportFolders[$folder]) {
        $subPath = Join-Path $path $subFolder
        New-Item -ItemType Directory -Force -Path $subPath | Out-Null
    }
}

# Function to export data with timestamp
function Export-ReportData {
    param(
        $Data,
        $Category,
        $SubCategory,
        $Name
    )
    
    try {
        # Create base reports directory if it doesn't exist
        $baseDir = "Tenantinfo\Reports"
        if (-not (Test-Path $baseDir)) {
            New-Item -ItemType Directory -Path $baseDir -Force | Out-Null
        }

        # Create category directory
        $categoryPath = Join-Path $baseDir $Category
        if (-not (Test-Path $categoryPath)) {
            New-Item -ItemType Directory -Path $categoryPath -Force | Out-Null
        }

        # Create subcategory directory
        $subCategoryPath = Join-Path $categoryPath $SubCategory
        if (-not (Test-Path $subCategoryPath)) {
            New-Item -ItemType Directory -Path $subCategoryPath -Force | Out-Null
        }

        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $fileName = "${Name}_${timestamp}.csv"
        $path = Join-Path $subCategoryPath $fileName

        # Check if data exists
        if ($null -eq $Data -or ($Data -is [Array] -and $Data.Count -eq 0)) {
            Write-Warning "No data available for $Category/$SubCategory/$Name"
            
            # Create empty CSV with headers to maintain structure
            [PSCustomObject]@{
                ReportName    = $Name
                Category      = $Category
                SubCategory   = $SubCategory
                GeneratedDate = Get-Date
                Status        = "No Data Available"
            } | Export-Csv -Path $path -NoTypeInformation
        }
        else {
            $Data | Export-Csv -Path $path -NoTypeInformation
            Write-Host "Exported $Category/$SubCategory/$Name to $path"
        }
    }
    catch {
        Write-Error "Failed to export"
    }
}
# Function to get access token
function Get-GraphToken {
    $tokenUrl = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token"
    $body = @{
        client_id     = $clientId
        scope         = "https://graph.microsoft.com/.default"
        client_secret = $clientSecret
        grant_type    = "client_credentials"
    }

    try {
        $response = Invoke-RestMethod -Method Post -Uri $tokenUrl -Body $body -ContentType "application/x-www-form-urlencoded"
        return $response.access_token
    }
    catch {
        Write-Error "Failed to get access token: $_"
        exit
    }
}

# Function to make Graph API calls
function Invoke-GraphRequest {
    param (
        [string]$Method = "GET",
        [string]$Endpoint,
        [string]$accessToken,
        [int]$MaxRetries = 3,
        [int]$RetryDelaySeconds = 5
    )

    $attempt = 0
    $success = $false
    $result = $null

    while (-not $success -and $attempt -lt $MaxRetries) {
        try {
            $headers = @{
                "Authorization" = "Bearer $accessToken"
                "Content-Type"  = "application/json"
                "ConsistencyLevel" = "eventual"
            }

            $response = Invoke-RestMethod -Method $Method `
                -Uri "https://graph.microsoft.com/v1.0/$Endpoint" `
                -Headers $headers

            # Handle paging
            $results = @()
            if ($response.value) {
                $results += $response.value
            }

            while ($response.'@odata.nextLink') {
                Start-Sleep -Seconds 1  # Avoid throttling
                $response = Invoke-RestMethod -Method $Method `
                    -Uri $response.'@odata.nextLink' `
                    -Headers $headers
                $results += $response.value
            }

            $success = $true
            $result = $results
        }
        catch {
            $attempt++
            
            if ($_.Exception.Response.StatusCode -eq 401) {
                Write-Host "Token expired, refreshing..."
                $accessToken = Get-GraphToken -TenantId $tenantId -ClientId $clientId -ClientSecret $clientSecret
                if (-not $accessToken) {
                    Write-Error "Failed to refresh token"
                    return $null
                }
            }
            elseif ($attempt -lt $MaxRetries) {
                Write-Host "Attempt $attempt failed. Retrying in $RetryDelaySeconds seconds..."
                Start-Sleep -Seconds $RetryDelaySeconds
            }
            else {
                Write-Warning "Error calling $Endpoint after $MaxRetries attempts: $_"
                return $null
            }
        }
    }

    return $result
}
Write-Host "Initializing Graph API connection..."
$accessToken = Get-GraphToken

if (-not $accessToken) {
    Write-Error "Failed to obtain valid access token. Exiting script."
    exit
}
# Function to format dates
function Format-GraphDate {
    param($date)
    if ($date) {
        return [DateTime]::Parse($date).ToString("yyyy-MM-dd HH:mm:ss")
    }
    return "Never"
}
function Get-UserProperties {
    param (
        [string]$accessToken
    )

    Write-Host "Getting detailed user information..."

    # Define properties to retrieve
    $select = @(
        "id", "userPrincipalName", "displayName", "givenName", "surname", 
        "mail", "otherMails", "userType", "accountEnabled", "ageGroup",
        "city", "companyName", "country", "createdDateTime", "department",
        "employeeId", "employeeType", "jobTitle", "lastPasswordChangeDateTime",
        "mobilePhone", "officeLocation", "onPremisesDistinguishedName",
        "onPremisesLastSyncDateTime", "onPremisesSamAccountName",
        "onPremisesSecurityIdentifier", "postalCode", "preferredLanguage",
        "state", "streetAddress", "usageLocation"
    ) -join ","

    # Get basic user information
    $users = Invoke-GraphRequest -AccessToken $accessToken -Endpoint "users?`$select=$select"

    if (-not $users) {
        Write-Warning "No users found or error retrieving users"
        return $null
    }

    $detailedUsers = @()
    $totalUsers = ($users | Measure-Object).Count
    $current = 0

    foreach ($user in $users) {
        $current++
        Write-Progress -Activity "Processing Users" -Status "Processing user $current of $totalUsers" `
            -PercentComplete (($current / $totalUsers) * 100)

        Write-Host "Processing user: $($user.userPrincipalName)"
        
        try {
            # Get additional user information in parallel
            $userDetails = @{
                Licenses       = Invoke-GraphRequest -AccessToken $accessToken -Endpoint "users/$($user.id)/licenseDetails"
                Groups         = Invoke-GraphRequest -AccessToken $accessToken -Endpoint "users/$($user.id)/memberOf"
                AuthMethods    = Invoke-GraphRequest -AccessToken $accessToken -Endpoint "users/$($user.id)/authentication/methods"
                Manager        = Invoke-GraphRequest -AccessToken $accessToken -Endpoint "users/$($user.id)?`$expand=manager(`$select=id,displayName)"
                DirectReports  = Invoke-GraphRequest -AccessToken $accessToken -Endpoint "users/$($user.id)/directReports"
                SignInActivity = Invoke-GraphRequest -AccessToken $accessToken -Endpoint "users/$($user.id)?`$select=signInActivity"
            }

            # Create user object with all properties
            $userObject = [PSCustomObject]@{
                # Basic Properties
                UserId                   = $user.id
                UserPrincipalName        = $user.userPrincipalName
                DisplayName              = $user.displayName
                GivenName                = $user.givenName
                Surname                  = $user.surname
                Mail                     = $user.mail
                OtherMails               = $user.otherMails -join ";"
                UserType                 = $user.userType
                AccountEnabled           = $user.accountEnabled
                
                # Profile Information
                Department               = $user.department
                JobTitle                 = $user.jobTitle
                EmployeeId               = $user.employeeId
                EmployeeType             = $user.employeeType
                CompanyName              = $user.companyName
                
                # Contact Information
                MobilePhone              = $user.mobilePhone
                OfficeLocation           = $user.officeLocation
                City                     = $user.city
                State                    = $user.state
                Country                  = $user.country
                PostalCode               = $user.postalCode
                StreetAddress            = $user.streetAddress
                
                # System Information
                CreatedDateTime          = $user.createdDateTime
                LastPasswordChange       = $user.lastPasswordChangeDateTime
                PreferredLanguage        = $user.preferredLanguage
                UsageLocation            = $user.usageLocation
                
                # On-Premises Information
                OnPremisesSamAccountName = $user.onPremisesSamAccountName
                OnPremisesSID            = $user.onPremisesSecurityIdentifier
                LastSyncDateTime         = $user.onPremisesLastSyncDateTime
                DistinguishedName        = $user.onPremisesDistinguishedName
                
                # License Information
                AssignedLicenses         = ($userDetails.Licenses | Select-Object -ExpandProperty skuPartNumber) -join ";"
                LicenseCount             = ($userDetails.Licenses | Measure-Object).Count
                
                # Group Information
                GroupMemberships         = ($userDetails.Groups | Select-Object -ExpandProperty displayName) -join ";"
                GroupCount               = ($userDetails.Groups | Measure-Object).Count
                SecurityGroups           = ($userDetails.Groups | Where-Object { $_.securityEnabled } | 
                    Select-Object -ExpandProperty displayName) -join ";"
                DistributionGroups       = ($userDetails.Groups | Where-Object { $_.mailEnabled } | 
                    Select-Object -ExpandProperty displayName) -join ";"
                
                # Authentication Information
                HasMFA                   = ($userDetails.AuthMethods.Count -gt 1)
                AuthenticationMethods    = ($userDetails.AuthMethods | ForEach-Object { $_.'@odata.type' }) -join ";"
                
                # Organizational Information
                ManagerName              = $userDetails.Manager.displayName
                DirectReportCount        = ($userDetails.DirectReports | Measure-Object).Count
                DirectReports            = ($userDetails.DirectReports | Select-Object -ExpandProperty displayName) -join ";"
                
                # Activity Information
                LastSignIn               = $userDetails.SignInActivity.lastSignInDateTime
                LastNonInteractiveSignIn = $userDetails.SignInActivity.lastNonInteractiveSignInDateTime
            }

            $detailedUsers += $userObject
        }
        catch {
            Write-Warning "Error processing user $($user.userPrincipalName): $_"
        }
    }

    Write-Progress -Activity "Processing Users" -Completed

    return $detailedUsers
}

function Get-SharePointSiteProperties {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$accessToken,
        [int]$MaxRetries = 3,
        [int]$RetryDelaySeconds = 5,
        [int]$PageSize = 100
    )

    Write-Host "Getting SharePoint sites..."
    $detailedSites = @()
    $processedCount = 0
    $errorCount = 0
    $startTime = Get-Date

    # Verified properties from Microsoft Graph API v1.0
    $baseProperties = @(
        "id",
        "createdDateTime",
        "description",
        "displayName",
        "lastModifiedDateTime",
        "name",
        "webUrl",
        "siteCollection"
    ) -join ","

    try {
        # Get initial sites list
        $sites = Invoke-GraphRequest -AccessToken $accessToken `
            -Endpoint "sites?`$select=$baseProperties&`$orderby=displayName" `
            -MaxRetries $MaxRetries `
            -PageSize $PageSize

        if (-not $sites) {
            Write-Warning "No SharePoint sites found or error retrieving sites"
            return $null
        }

        $totalSites = ($sites | Measure-Object).Count
        Write-Host "Found $totalSites sites to process"

        foreach ($site in $sites) {
            $processedCount++
            $siteStartTime = Get-Date

            Write-Progress -Activity "Processing SharePoint Sites" `
                -Status "Processing site $processedCount of $totalSites" `
                -PercentComplete (($processedCount / $totalSites) * 100) `
                -CurrentOperation $site.displayName

            try {
                Write-Host "Processing site ($processedCount/$totalSites): $($site.displayName)"

                # Get detailed site information
                $siteDetails = $null
                try {
                    $siteDetails = Invoke-GraphRequest -AccessToken $accessToken `
                        -Endpoint "sites/$($site.id)" `
                        -MaxRetries $MaxRetries
                }
                catch {
                    Write-Warning "Error getting details for site $($site.displayName): $_"
                }

                # Get drives (document libraries)
                $drives = @()
                try {
                    $drives = Invoke-GraphRequest -AccessToken $accessToken `
                        -Endpoint "sites/$($site.id)/drives?`$select=id,name,driveType,quota,lastModifiedDateTime,webUrl,createdDateTime" `
                        -PageSize $PageSize
                }
                catch {
                    Write-Warning "Error getting drives for site $($site.displayName): $_"
                }

                # Get lists
                $lists = @()
                try {
                    $lists = Invoke-GraphRequest -AccessToken $accessToken `
                        -Endpoint "sites/$($site.id)/lists?`$select=id,displayName,createdDateTime,lastModifiedDateTime,list" `
                        -PageSize $PageSize
                }
                catch {
                    Write-Warning "Error getting lists for site $($site.displayName): $_"
                }

                # Calculate storage metrics
                $storageMetrics = @{
                    UsedGB = 0
                    TotalGB = 0
                    PercentageUsed = 0
                }

                if ($drives) {
                    $totalUsed = 0
                    $totalQuota = 0

                    foreach ($drive in $drives) {
                        if ($drive.quota) {
                            $totalUsed += if ($drive.quota.used) { $drive.quota.used } else { 0 }
                            $totalQuota += if ($drive.quota.total) { $drive.quota.total } else { 0 }
                        }
                    }

                    $storageMetrics.UsedGB = [math]::Round($totalUsed / 1GB, 2)
                    $storageMetrics.TotalGB = [math]::Round($totalQuota / 1GB, 2)
                    $storageMetrics.PercentageUsed = if ($totalQuota -gt 0) {
                        [math]::Round(($totalUsed / $totalQuota) * 100, 2)
                    } else { 0 }
                }

                # Create site object with all properties
                $siteObject = [PSCustomObject]@{
                    # Basic Properties
                    Id = $site.id
                    DisplayName = $site.displayName
                    Name = $site.name
                    Description = $site.description
                    WebUrl = $site.webUrl
                    CreatedDateTime = $site.createdDateTime
                    LastModifiedDateTime = $site.lastModifiedDateTime

                    # Site Collection Information
                    SiteCollectionId = $site.siteCollection.id
                    SiteCollectionHostname = $site.siteCollection.hostname
                    RootSite = $site.siteCollection.root

                    # Storage Information
                    StorageUsedGB = $storageMetrics.UsedGB
                    StorageTotalGB = $storageMetrics.TotalGB
                    StoragePercentageUsed = $storageMetrics.PercentageUsed

                    # Document Libraries (Drives)
                    DriveCount = ($drives | Measure-Object).Count
                    Drives = $drives | Select-Object id, name, driveType, webUrl, lastModifiedDateTime, createdDateTime, @{
                        Name = 'StorageUsedGB'
                        Expression = { if ($_.quota.used) { [math]::Round($_.quota.used / 1GB, 2) } else { 0 } }
                    }

                    # Lists
                    ListCount = ($lists | Measure-Object).Count
                    Lists = $lists | Select-Object id, displayName, createdDateTime, lastModifiedDateTime

                    # Site Type
                    IsTeamsSite = $site.webUrl -like "*/teams*"
                    IsCommunicationSite = $null -ne ($lists | Where-Object { $_.list.template -eq "CommsSite" })
                    IsHubSite = $siteDetails.isHubSite

                    # Status Information
                    ProcessingStatus = "Success"
                    ProcessingDuration = [math]::Round(((Get-Date) - $siteStartTime).TotalSeconds, 2)
                    LastProcessed = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                    ErrorMessage = $null
                }

                $detailedSites += $siteObject

            }
            catch {
                $errorCount++
                $errorMessage = "Error processing site $($site.displayName): $_"
                Write-Warning $errorMessage
                
                $detailedSites += [PSCustomObject]@{
                    Id = $site.id
                    DisplayName = $site.displayName
                    WebUrl = $site.webUrl
                    ProcessingStatus = "Failed"
                    ProcessingDuration = [math]::Round(((Get-Date) - $siteStartTime).TotalSeconds, 2)
                    LastProcessed = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                    ErrorMessage = $errorMessage
                }
            }
        }

        Write-Progress -Activity "Processing SharePoint Sites" -Completed

        # Create summary
        $summary = [PSCustomObject]@{
            TotalSites = $totalSites
            ProcessedSites = $processedCount
            SuccessfulSites = ($detailedSites | Where-Object { $_.ProcessingStatus -eq "Success" }).Count
            FailedSites = $errorCount
            TotalStorageGB = ($detailedSites | Measure-Object -Property StorageUsedGB -Sum).Sum
            AverageStorageGB = ($detailedSites | Where-Object { $_.StorageUsedGB -gt 0 } | 
                Measure-Object -Property StorageUsedGB -Average).Average
            TotalDrives = ($detailedSites | Measure-Object -Property DriveCount -Sum).Sum
            TotalLists = ($detailedSites | Measure-Object -Property ListCount -Sum).Sum
            TeamsSites = ($detailedSites | Where-Object { $_.IsTeamsSite } | Measure-Object).Count
            CommunicationSites = ($detailedSites | Where-Object { $_.IsCommunicationSite } | Measure-Object).Count
            HubSites = ($detailedSites | Where-Object { $_.IsHubSite } | Measure-Object).Count
            ProcessingDuration = [math]::Round(((Get-Date) - $startTime).TotalMinutes, 2)
            CompletedDateTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        }

        # Export the data
        Export-ReportData -Data $detailedSites -Category "SharePoint" -SubCategory "Sites" -Name "DetailedSiteProperties"
        Export-ReportData -Data $summary -Category "SharePoint" -SubCategory "Summary" -Name "SiteSummary"

        # Final status message
        Write-Host "SharePoint site processing completed in $($summary.ProcessingDuration) minutes."
        Write-Host "Successful: $($summary.SuccessfulSites), Failed: $($summary.FailedSites)"
        Write-Host "Total Storage Used: $($summary.TotalStorageGB) GB"

        return $detailedSites
    }
    catch {
        Write-Error "Fatal error during SharePoint site processing: $_"
        return $null
    }
}


$sharePointSites = Get-SharePointSiteProperties -AccessToken $accessToken

if ($sharePointSites) {
    Write-Host "Successfully retrieved properties for $($sharePointSites.Count) SharePoint sites"
}
else {
    Write-Warning "No SharePoint sites were retrieved"
}


Write-Host "Getting user properties..."
$userProperties = Get-UserProperties -AccessToken $accessToken

# Export user properties
if ($userProperties) {
    Export-ReportData -Data $userProperties -Category "Users" -SubCategory "Properties" -Name "DetailedUserProperties"
}

# Get Users Information
Write-Host "Getting users information..."
$users = Invoke-GraphRequest -AccessToken $accessToken -Endpoint "users?`$select=id,userPrincipalName,displayName,userType,accountEnabled,mail,assignedLicenses"
Export-ReportData -Data $users -Category "Users" -SubCategory "BasicInfo" -Name "AllUsers"

# Get Guest Users
Write-Host "Processing guest users..."
$guestUsers = $users | Where-Object { $_.userType -eq 'Guest' } | Select-Object -ErrorAction SilentlyContinue

if ($null -eq $guestUsers) {
    Write-Host "No guest users found in the tenant"
    $guestUsers = @() # Ensure we have an empty array rather than null
}

Export-ReportData -Data $guestUsers -Category "Users" -SubCategory "GuestAccounts" -Name "GuestUsers"


# Get Sign-in Activity and Inactive Users
Write-Host "Getting sign-in activity..."
$signInLogs = Invoke-GraphRequest -AccessToken $accessToken -Endpoint "auditLogs/signIns?`$top=1000"
$inactiveThreshold = (Get-Date).AddDays(-90)
$inactiveUsers = $users | Where-Object {
    $lastSignIn = $signInLogs | 
    Where-Object { $_.userPrincipalName -eq $_.userPrincipalName } |
    Sort-Object createdDateTime -Descending |
    Select-Object -First 1
    (!$lastSignIn) -or ([DateTime]$lastSignIn.createdDateTime -lt $inactiveThreshold)
}
Export-ReportData -Data $inactiveUsers -Category "Users" -SubCategory "InactiveAccounts" -Name "InactiveUsers"

# Get Product Usage Information
Write-Host "Getting product usage information..."
$productUsage = @()
foreach ($user in $users) {
    Write-Host "Getting usage for user: $($user.userPrincipalName)"
    
    # Define service endpoints and products
    $serviceEndpoints = @{
        "Office"     = "reports/getOffice365ActiveUserDetail(period='D90')"
        "Email"      = "reports/getEmailActivityUserDetail(period='D90')"
        "OneDrive"   = "reports/getOneDriveActivityUserDetail(period='D90')"
        "SharePoint" = "reports/getSharePointActivityUserDetail(period='D90')"
        "Teams"      = "reports/getTeamsUserActivityUserDetail(period='D90')"
        "Yammer"     = "reports/getYammerActivityUserDetail(period='D90')"
        "Copilot"     = "reports/getOffice365ActiveUserDetail(period='D90')"

    }

    # Define specific products to track
    $productsToTrack = @{
        "VISIO"          = @("VISIOCLIENT", "VISIOPRO", "VISIOSTD")
        "PROJECT"        = @("PROJECTCLIENT", "PROJECTPREMIUM", "PROJECTPROFESSIONAL")
        "POWER_PLATFORM" = @("POWER_BI_PRO", "POWER_BI_PREMIUM", "FLOW_FREE", "POWERAPPS_VIRAL")
        "WINDOWS"        = @("WIN10_PRO_ENT_SUB", "WIN_DEF_ATP", "WINBIZ")
        "DYNAMICS"       = @("DYN365_ENTERPRISE_PLAN1", "DYN365_ENTERPRISE_CUSTOMER", "DYN365_FINANCIALS_BUSINESS")
        "AZURE"          = @("AZUREAD_PREMIUM", "AZUREACTIVEDIRECTORY", "ATA")
    }

    $userUsage = [PSCustomObject]@{
        UserPrincipalName = $user.userPrincipalName
        DisplayName       = $user.displayName
    }

    # Get standard service usage
    foreach ($service in $serviceEndpoints.Keys) {
        try {
            $reportData = Invoke-GraphRequest -AccessToken $accessToken `
                -Endpoint "$($serviceEndpoints[$service])"
            
            if ($reportData) {
                $lastActivity = $reportData | 
                Where-Object { $_.lastActivityDate } | 
                Sort-Object lastActivityDate -Descending | 
                Select-Object -First 1
                
                $userUsage | Add-Member -NotePropertyName "Last${service}Usage" -NotePropertyValue ($lastActivity.lastActivityDate ?? "Never")
                $userUsage | Add-Member -NotePropertyName "${service}IsActive" -NotePropertyValue ($lastActivity.lastActivityDate -ne $null)
            }
            else {
                $userUsage | Add-Member -NotePropertyName "Last${service}Usage" -NotePropertyValue "Never"
                $userUsage | Add-Member -NotePropertyName "${service}IsActive" -NotePropertyValue $false
            }
        }
        catch {
            Write-Warning "Error getting $service usage for $($user.userPrincipalName): $_"
            $userUsage | Add-Member -NotePropertyName "Last${service}Usage" -NotePropertyValue "Error"
            $userUsage | Add-Member -NotePropertyName "${service}IsActive" -NotePropertyValue $false
        }
    }

    # Get license details for additional products
    try {
        $licenseDetails = Invoke-GraphRequest -AccessToken $accessToken -Endpoint "users/$($user.id)/licenseDetails"
        
        foreach ($productCategory in $productsToTrack.Keys) {
            $hasProduct = $false
            $productStatus = @()

            foreach ($license in $licenseDetails) {
                $matchingPlans = $license.servicePlans | 
                Where-Object { $productsToTrack[$productCategory] -contains $_.servicePlanName }

                if ($matchingPlans) {
                    $hasProduct = $true
                    foreach ($plan in $matchingPlans) {
                        $productStatus += [PSCustomObject]@{
                            Name     = $plan.servicePlanName
                            Status   = $plan.provisioningStatus
                            LastSeen = $license.skuPartNumber
                        }
                    }
                }
            }

            $userUsage | Add-Member -NotePropertyName "Has${productCategory}" -NotePropertyValue $hasProduct
            $userUsage | Add-Member -NotePropertyName "${productCategory}_Details" -NotePropertyValue ($productStatus | ConvertTo-Json -Compress)
        }
    }
    catch {
        Write-Warning "Error getting license details for $($user.userPrincipalName): $_"
        foreach ($productCategory in $productsToTrack.Keys) {
            $userUsage | Add-Member -NotePropertyName "Has${productCategory}" -NotePropertyValue $false
            $userUsage | Add-Member -NotePropertyName "${productCategory}_Details" -NotePropertyValue "{}"
        }
    }

    # Get activation data
    try {
        $activationData = Invoke-GraphRequest -AccessToken $accessToken -Endpoint "reports/getOffice365ActivationsUserDetail"
        $userActivations = $activationData | Where-Object { $_.userPrincipalName -eq $user.userPrincipalName }
        
        $userUsage | Add-Member -NotePropertyName "OfficeActivations" -NotePropertyValue ($userActivations.Count ?? 0)
        $userUsage | Add-Member -NotePropertyName "LastActivation" -NotePropertyValue ($userActivations | 
            Sort-Object activationDate -Descending | 
            Select-Object -First 1 -ExpandProperty activationDate)

        # Track specific product activations
        $activatedProducts = $userActivations | Group-Object productType
        $userUsage | Add-Member -NotePropertyName "ActivatedProducts" -NotePropertyValue ($activatedProducts.Name -join ";")
        
        foreach ($product in $activatedProducts) {
            $productName = $product.Name -replace '[^a-zA-Z0-9]', ''
            $userUsage | Add-Member -NotePropertyName "${productName}_ActivationCount" -NotePropertyValue $product.Count
            $userUsage | Add-Member -NotePropertyName "${productName}_LastActivation" -NotePropertyValue (
                $product.Group | Sort-Object activationDate -Descending | 
                Select-Object -First 1 -ExpandProperty activationDate
            )
        }
    }
    catch {
        Write-Warning "Error getting activation data for $($user.userPrincipalName): $_"
        $userUsage | Add-Member -NotePropertyName "OfficeActivations" -NotePropertyValue 0
        $userUsage | Add-Member -NotePropertyName "LastActivation" -NotePropertyValue $null
        $userUsage | Add-Member -NotePropertyName "ActivatedProducts" -NotePropertyValue ""
    }

    $productUsage += $userUsage
}

# Create detailed license summary
$licenseSummary = @()
foreach ($sku in $subscriptions) {
    $usageData = $licenseHistory | Where-Object { $_.LicenseName -eq $sku.skuPartNumber }
    
    $licenseSummary += [PSCustomObject]@{
        LicenseName       = $sku.skuPartNumber
        TotalAssigned     = ($usageData | Measure-Object).Count
        ActivelyUsed      = ($usageData | Where-Object { $_.LastUsedDate -ne "Never" } | Measure-Object).Count
        NeverUsed         = ($usageData | Where-Object { $_.LastUsedDate -eq "Never" } | Measure-Object).Count
        AvailableLicenses = $sku.prepaidUnits.enabled - $sku.consumedUnits
        ServicePlans      = ($sku.servicePlans | ConvertTo-Json -Compress)
        Category          = switch -Wildcard ($sku.skuPartNumber) {
            "VISIO*" { "Visio" }
            "PROJECT*" { "Project" }
            "POWER_*" { "Power Platform" }
            "DYN365*" { "Dynamics" }
            "AZURE*" { "Azure" }
            "WIN*" { "Windows" }
            default { "Other" }
        }
    }
}

Export-ReportData -Data $productUsage -Category "Users" -SubCategory "ProductUsage" -Name "DetailedProductUsage"
Export-ReportData -Data $licenseSummary -Category "Licenses" -SubCategory "Summary" -Name "DetailedLicenseSummary"


# Get Teams Information
Write-Host "Getting Teams information..."
$teams = Invoke-GraphRequest -AccessToken $accessToken -Endpoint "teams?`$select=id,displayName,description,visibility"
Export-ReportData -Data $teams -Category "Teams" -SubCategory "BasicInfo" -Name "AllTeams"

# Get Teams with Guests
Write-Host "Processing Teams with guests..."
$teamsWithGuests = foreach ($team in $teams) {
    $members = Invoke-GraphRequest -AccessToken $accessToken -Endpoint "teams/$($team.id)/members"
    $guestCount = ($members | Where-Object { $_.'@odata.type' -eq '#microsoft.graph.aadUserConversationMember' -and $_.roles -contains 'guest' } | Measure-Object).Count
    if ($guestCount -gt 0) {
        [PSCustomObject]@{
            TeamName   = $team.displayName
            GuestCount = $guestCount
        }
    }
}
Export-ReportData -Data $teamsWithGuests -Category "Teams" -SubCategory "GuestAccess" -Name "TeamsWithGuests"

# Get Teams Channel Usage
Write-Host "Getting Teams channel usage..."
$teamsChannels = @()
foreach ($team in $teams) {
    $channels = Invoke-GraphRequest -AccessToken $accessToken -Endpoint "teams/$($team.id)/channels"
    $teamsChannels += [PSCustomObject]@{
        TeamName       = $team.displayName
        ChannelCount   = $channels.Count
        DefaultChannel = ($channels | Where-Object { $_.displayName -eq "General" }).id
        CustomChannels = ($channels | Where-Object { $_.displayName -ne "General" }).displayName -join ';'
    }
}
Export-ReportData -Data $teamsChannels -Category "Teams" -SubCategory "Channels" -Name "TeamsChannels"

# Get SharePoint Sites
Write-Host "Getting SharePoint sites..."
$sites = Invoke-GraphRequest -AccessToken $accessToken -Endpoint "sites?search=*"
Export-ReportData -Data $sites -Category "SharePoint" -SubCategory "Sites" -Name "AllSites"

# Get SharePoint Storage Usage
Write-Host "Getting SharePoint storage usage..."
$siteStorage = @()
foreach ($site in $sites) {
    try {
        $storage = Invoke-GraphRequest -AccessToken $accessToken -Endpoint "sites/$($site.id)/drive"
        
        # Check if storage data exists and quota is not zero
        if ($storage -and $storage.quota -and $storage.quota.total -gt 0) {
            $siteStorage += [PSCustomObject]@{
                SiteName       = $site.displayName
                Url            = $site.webUrl
                StorageUsedGB  = [math]::Round($storage.quota.used / 1GB, 2)
                StorageTotalGB = [math]::Round($storage.quota.total / 1GB, 2)
                PercentageUsed = [math]::Round(($storage.quota.used / $storage.quota.total) * 100, 2)
                LastModified   = $site.lastModifiedDateTime
            }
        }
        else {
            # Add site with zero or unknown storage
            $siteStorage += [PSCustomObject]@{
                SiteName       = $site.displayName
                Url            = $site.webUrl
                StorageUsedGB  = 0
                StorageTotalGB = 0
                PercentageUsed = 0
                LastModified   = $site.lastModifiedDateTime
            }
        }
    }
    catch {
        Write-Warning "Failed to get storage info for site $($site.displayName): $_"
        # Add site with error status
        $siteStorage += [PSCustomObject]@{
            SiteName       = $site.displayName
            Url            = $site.webUrl
            StorageUsedGB  = 0
            StorageTotalGB = 0
            PercentageUsed = 0
            LastModified   = $site.lastModifiedDateTime
            Error          = $_.Exception.Message
        }
    }
}

Export-ReportData -Data $siteStorage -Category "SharePoint" -SubCategory "Storage" -Name "StorageUsage"


# Get Security Score
Write-Host "Getting security score..."
$securityScore = Invoke-GraphRequest -AccessToken $accessToken -Endpoint "security/secureScores?`$top=1"
$securityScoreReport = [PSCustomObject]@{
    CurrentScore    = $securityScore[0].currentScore
    MaxScore        = $securityScore[0].maxScore
    PercentageScore = [math]::Round(($securityScore[0].currentScore / $securityScore[0].maxScore) * 100, 2)
    LastUpdated     = Format-GraphDate $securityScore[0].createdDateTime
}
Export-ReportData -Data $securityScoreReport -Category "Security" -SubCategory "Score" -Name "SecurityScore"

# Get MFA Status
Write-Host "Getting MFA status..."
$mfaStatus = @()
foreach ($user in $users) {
    $authMethods = Invoke-GraphRequest -AccessToken $accessToken -Endpoint "users/$($user.id)/authentication/methods"
    $mfaStatus += [PSCustomObject]@{
        UserPrincipalName = $user.userPrincipalName
        DisplayName       = $user.displayName
        HasMFA            = ($authMethods.Count -gt 1)
        AuthMethods       = ($authMethods | ForEach-Object { $_.'@odata.type' }) -join ';'
        IsAdmin           = ($user.assignedRoles.Count -gt 0)
    }
}
Export-ReportData -Data $mfaStatus -Category "Users" -SubCategory "MFAStatus" -Name "MFAStatus"

# Get Device Compliance
Write-Host "Getting device compliance..."
$deviceCompliance = Invoke-GraphRequest -AccessToken $accessToken -Endpoint "deviceManagement/managedDevices?`$select=id,deviceName,userPrincipalName,complianceState,lastSyncDateTime,osVersion,model"
Export-ReportData -Data $deviceCompliance -Category "Security" -SubCategory "DeviceCompliance" -Name "DeviceStatus"

# Get License Information
Write-Host "Getting detailed license information..."
$licenseHistory = @()
$subscriptions = Invoke-GraphRequest -AccessToken $accessToken -Endpoint "subscribedSkus"

foreach ($user in $users) {
    Write-Host "Getting license history for user: $($user.userPrincipalName)"
    
    # Get current licenses
    $currentLicenses = Invoke-GraphRequest -AccessToken $accessToken -Endpoint "users/$($user.id)/licenseDetails"
    
    # Get audit logs for license changes
    $auditFilter = "targetResources/any(t:t/resourceId eq '$($user.id)') and activityDisplayName eq 'Add user' or activityDisplayName eq 'Update user' or activityDisplayName eq 'Change user license'"
    $licenseAudits = Invoke-GraphRequest -AccessToken $accessToken -Endpoint "auditLogs/directoryAudits?`$filter=$auditFilter"

    # Get usage reports
    $usageReports = Invoke-GraphRequest -AccessToken $accessToken -Endpoint "reports/getOffice365ActivationsUserDetail"
    $userUsage = $usageReports | Where-Object { $_.userPrincipalName -eq $user.userPrincipalName }

    foreach ($license in $currentLicenses) {
        $lastUsed = "Never"
        $activationStatus = "Not Activated"

        $activation = $userUsage | Where-Object { $_.productType -eq $license.skuPartNumber }
        if ($activation) {
            $lastUsed = $activation.lastActivatedDate
            $activationStatus = $activation.activationStatus
        }

        $assignmentDate = ($licenseAudits | 
            Where-Object { $_.targetResources.modifiedProperties.displayName -contains "AssignedLicenses" } |
            Select-Object -First 1).activityDateTime

        $licenseHistory += [PSCustomObject]@{
            UserPrincipalName = $user.userPrincipalName
            DisplayName       = $user.displayName
            LicenseName       = $license.skuPartNumber
            AssignmentDate    = $assignmentDate
            LastUsedDate      = $lastUsed
            ActivationStatus  = $activationStatus
            EnabledServices   = ($license.servicePlans | Where-Object { $_.provisioningStatus -eq "Success" }).servicePlanName -join ';'
            DisabledServices  = ($license.servicePlans | Where-Object { $_.provisioningStatus -ne "Success" }).servicePlanName -join ';'
        }
    }
}

Export-ReportData -Data $licenseHistory -Category "Licenses" -SubCategory "History" -Name "LicenseHistory"

# Create License Usage Summary
$licenseUsageSummary = @()
foreach ($sku in $subscriptions) {
    $usageData = $licenseHistory | Where-Object { $_.LicenseName -eq $sku.skuPartNumber }
    
    $licenseUsageSummary += [PSCustomObject]@{
        LicenseName       = $sku.skuPartNumber
        TotalAssigned     = ($usageData | Measure-Object).Count
        ActivelyUsed      = ($usageData | Where-Object { $_.LastUsedDate -ne "Never" } | Measure-Object).Count
        NeverUsed         = ($usageData | Where-Object { $_.LastUsedDate -eq "Never" } | Measure-Object).Count
        AvailableLicenses = $sku.prepaidUnits.enabled - $sku.consumedUnits
        WastedLicenses    = ($usageData | Where-Object { $_.LastUsedDate -eq "Never" } | Measure-Object).Count
    }
}

Export-ReportData -Data $licenseUsageSummary -Category "Licenses" -SubCategory "Usage" -Name "LicenseUsage"

# Create Enhanced Summary Report
$enhancedSummary = [PSCustomObject]@{
    ReportGenerated           = Get-Date
    TotalUsers                = ($users | Measure-Object).Count
    GuestUsers                = ($guestUsers | Measure-Object).Count
    InactiveUsers             = ($inactiveUsers | Measure-Object).Count
    TotalTeams                = ($teams | Measure-Object).Count
    TeamsWithGuests           = ($teamsWithGuests | Measure-Object).Count
    TotalSites                = ($sites | Measure-Object).Count
    SecurityScore             = $securityScoreReport.PercentageScore
    UsersWithMFA              = ($mfaStatus | Where-Object { $_.HasMFA } | Measure-Object).Count
    AdminsWithoutMFA          = ($mfaStatus | Where-Object { $_.IsAdmin -and -not $_.HasMFA } | Measure-Object).Count
    NonCompliantDevices       = ($deviceCompliance | Where-Object { $_.complianceState -ne "compliant" } | Measure-Object).Count
    TotalDevices              = ($deviceCompliance | Measure-Object).Count
    TotalSharePointStorageGB  = ($siteStorage | Measure-Object -Property StorageUsedGB -Sum).Sum
    UnusedLicenseCount        = ($licenseHistory | Where-Object { $_.LastUsedDate -eq "Never" } | Measure-Object).Count
    PartiallyUsedLicenseCount = ($licenseHistory | Where-Object { 
            $_.EnabledServices -and ($_.EnabledServices -split ';').Count -gt ($_.DisabledServices -split ';').Count 
        } | Measure-Object).Count
    LastReportGeneration      = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
}

Export-ReportData -Data $enhancedSummary -Category "Reports" -SubCategory "Summary" -Name "EnhancedSummary"

$dashboardMetrics = @{
    # User Metrics
    TotalUsers              = @($users).Count
    ActiveUsers             = @($users | Where-Object { $_.accountEnabled -eq $true }).Count
    InactiveUsers           = @($inactiveUsers).Count
    GuestUsers              = @($guestUsers).Count
    
    # Security Metrics
    SecurityScore           = [math]::Round($securityScoreReport.PercentageScore)
    CurrentSecurityScore    = [math]::Round($securityScoreReport.CurrentScore)
    MaxSecurityScore        = [math]::Round($securityScoreReport.MaxScore)
    UsersWithMFA            = @($mfaStatus | Where-Object { $_.HasMFA }).Count
    UsersWithoutMFA         = @($mfaStatus | Where-Object { -not $_.HasMFA }).Count
    AdminsWithoutMFA        = @($mfaStatus | Where-Object { $_.IsAdmin -and -not $_.HasMFA }).Count
    
    # License Metrics
    TotalLicenses           = @($subscriptions).Count
    UnusedLicenses          = @($licenseHistory | Where-Object { $_.LastUsedDate -eq "Never" }).Count
    PartialLicenses         = @($licenseHistory | Where-Object { 
            $_.EnabledServices -and ($_.EnabledServices -split ';').Count -gt ($_.DisabledServices -split ';').Count 
        }).Count
    
    # Teams Metrics
    TotalTeams              = @($teams).Count
    TeamsWithGuests         = @($teamsWithGuests).Count
    TeamsActiveUsers        = @($productUsage | Where-Object { $_.'TeamsIsActive' -eq $true }).Count
    
    # SharePoint Metrics
    TotalSites              = @($sites).Count
    TotalStorageGB          = [math]::Round(($siteStorage | Measure-Object -Property StorageUsedGB -Sum).Sum, 1)
    ActiveSites             = @($sharePointSites | Where-Object { $_.LastModifiedDateTime -gt (Get-Date).AddDays(-90) }).Count
    ExternalSharingSites    = @($sharePointSites | Where-Object { $_.ExternalSharingEnabled }).Count
    
    # Device Metrics
    TotalDevices            = @($deviceCompliance).Count
    NonCompliantDevices     = @($deviceCompliance | Where-Object { $_.complianceState -ne "compliant" }).Count
    CompliantDevices        = @($deviceCompliance | Where-Object { $_.complianceState -eq "compliant" }).Count
    
    # Activity Metrics
    ExchangeActiveUsers     = @($productUsage | Where-Object { $_.'EmailIsActive' -eq $true }).Count
    SharePointActiveUsers   = @($productUsage | Where-Object { $_.'SharePointIsActive' -eq $true }).Count
    OneDriveActiveUsers     = @($productUsage | Where-Object { $_.'OneDriveIsActive' -eq $true }).Count
    

    TotalStorageUsedGB      = [math]::Round(($siteStorage | Measure-Object -Property StorageUsedGB -Sum).Sum, 2)
    TotalStorageAvailableGB = [math]::Round(($siteStorage | Measure-Object -Property StorageTotalGB -Sum).Sum, 2)
    SiteDetails             = $siteStorage | Sort-Object -Property StorageUsedGB -Descending | ForEach-Object {
        @{
            SiteName       = $_.SiteName
            Url            = $_.Url
            UsedGB         = [math]::Round($_.StorageUsedGB, 2)
            TotalGB        = [math]::Round($_.StorageTotalGB, 2)
            PercentageUsed = [math]::Round($_.PercentageUsed, 2)
            LastModified   = $_.LastModified
        }
    }
    # License Distribution
    LicenseDetails          = $licenseUsageSummary | ForEach-Object {
        @{
            Name       = $_.LicenseName
            Total      = $_.TotalAssigned + $_.AvailableLicenses
            Used       = $_.TotalAssigned
            Active     = $_.ActivelyUsed
            Unused     = $_.NeverUsed
            Percentage = if (($_.TotalAssigned + $_.AvailableLicenses) -gt 0) {
                [math]::Round(($_.TotalAssigned / ($_.TotalAssigned + $_.AvailableLicenses)) * 100)
            }
            else { 0 }
        }
    }
}

# Generate HTML Report
$htmlReport = @"
<!DOCTYPE html>
<html lang="en" class="h-full bg-gray-50">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Microsoft 365 Tenant Dashboard</title>
    <link href="https://cdnjs.cloudflare.com/ajax/libs/tailwindcss/2.2.19/tailwind.min.css" rel="stylesheet">
    <link href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css" rel="stylesheet">
    <script src="https://cdnjs.cloudflare.com/ajax/libs/chart.js/3.7.0/chart.min.js"></script>
    <style>
        /* Custom Styling */
        .metric-card {
            transition: all 0.3s ease;
        }
        .metric-card:hover {
            transform: translateY(-5px);
            box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1), 0 2px 4px -1px rgba(0, 0, 0, 0.06);
        }
        .progress-ring {
            transform: rotate(-90deg);
        }
        .chart-container {
            position: relative;
            height: 300px;
            width: 100%;
        }
        @keyframes fadeIn {
            from { opacity: 0; }
            to { opacity: 1; }
        }
        .animate-fadeIn {
            animation: fadeIn 0.5s ease-in;
        }
        
        /* Responsive Adjustments */
        @media (max-width: 640px) {
            .chart-container {
                height: 200px;
            }
            .metric-grid {
                grid-template-columns: repeat(1, 1fr);
            }
        }
        
        /* Dark Mode Support */
        @media (prefers-color-scheme: dark) {
            .dark-mode-text {
                color: #e5e7eb;
            }
            .dark-mode-bg {
                background-color: #1f2937;
            }
        }

        /* Custom Scrollbar */
        ::-webkit-scrollbar {
            width: 8px;
            height: 8px;
        }
        ::-webkit-scrollbar-track {
            background: #f1f1f1;
        }
        ::-webkit-scrollbar-thumb {
            background: #888;
            border-radius: 4px;
        }
        ::-webkit-scrollbar-thumb:hover {
            background: #555;
        }

        /* Loading States */
        .loading {
            position: relative;
        }
        .loading:after {
            content: '';
            position: absolute;
            top: 0;
            left: 0;
            right: 0;
            bottom: 0;
            background: rgba(255, 255, 255, 0.7);
            display: flex;
            justify-content: center;
            align-items: center;
            font-size: 1.2em;
            color: #4B5563;
        }

        /* Print Styles */
        @media print {
            .no-print {
                display: none;
            }
            .page-break {
                page-break-before: always;
            }
        }

        /* Custom Card Styles */
        .stat-card {
            @apply bg-white rounded-lg shadow-sm p-6 transition-all duration-300;
        }
        .stat-card:hover {
            @apply shadow-lg;
        }
    </style>
</head>
<body class="h-full">
    <!-- Navigation -->
    <nav class="bg-white shadow-sm sticky top-0 z-50">
        <div class="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
            <div class="flex justify-between h-16">
                <div class="flex items-center">
                    <img class="h-8 w-auto" src="https://www.microsoft.com/favicon.ico" alt="Microsoft">
                    <span class="ml-2 text-xl font-semibold text-gray-900">M365 Tenant Dashboard</span>
                </div>
                <div class="flex items-center space-x-4">
                    <span class="text-sm text-gray-500">Last Updated: $($enhancedSummary.LastReportGeneration)</span>
                    <button class="no-print bg-blue-500 hover:bg-blue-600 text-white px-4 py-2 rounded-md text-sm font-medium transition-colors duration-200" onclick="window.print()">
                        Print Report
                    </button>
                </div>
            </div>
        </div>
    </nav>

    <!-- Main Content -->
    <main class="max-w-7xl mx-auto py-6 px-4 sm:px-6 lg:px-8">
        <!-- Quick Stats Grid -->
        <div class="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4 mb-8">
            <!-- Users Card -->
            <div class="stat-card">
                <div class="flex items-center justify-between">
                    <div>
                        <p class="text-gray-500 text-sm font-medium">Total Users</p>
                        <p class="text-2xl font-bold text-gray-900">$($dashboardMetrics.TotalUsers)</p>
                    </div>
                    <div class="rounded-full p-3 bg-blue-100">
                        <i class="fas fa-users text-blue-500 text-xl"></i>
                    </div>
                </div>
                <div class="mt-4">
                    <div class="flex justify-between text-sm">
                        <span class="text-green-500">Active: $($dashboardMetrics.ActiveUsers)</span>
                        <span class="text-red-500">Inactive: $($dashboardMetrics.InactiveUsers)</span>
                    </div>
                </div>
            </div>

            <!-- Security Score Card -->
            <div class="stat-card">
                <div class="flex items-center justify-between">
                    <div>
                        <p class="text-gray-500 text-sm font-medium">Security Score</p>
                        <p class="text-2xl font-bold text-gray-900">$($dashboardMetrics.SecurityScore)%</p>
                    </div>
                    <div class="relative">
                        <svg class="progress-ring" width="50" height="50">
                            <circle class="text-gray-200" stroke="currentColor" stroke-width="5" fill="transparent" r="20" cx="25" cy="25"/>
                            <circle class="text-blue-600" stroke="currentColor" stroke-width="5" fill="transparent" r="20" cx="25" cy="25"
                                stroke-dasharray="125.6"
                                stroke-dashoffset="$([math]::Round(125.6 - ($dashboardMetrics.SecurityScore / 100 * 125.6)))"/>
                        </svg>
                    </div>
                </div>
                <div class="mt-4">
                    <div class="text-sm text-red-500">
                        $($dashboardMetrics.AdminsWithoutMFA) Admins without MFA
                    </div>
                </div>
            </div>

            <!-- Storage Card -->
            <div class="stat-card">
                <div class="flex items-center justify-between">
                    <div>
                        <p class="text-gray-500 text-sm font-medium">Total Storage</p>
                        <p class="text-2xl font-bold text-gray-900">$($dashboardMetrics.StorageDetails.TotalStorageUsedGB) GB</p>
                    </div>
                    <div class="rounded-full p-3 bg-purple-100">
                        <i class="fas fa-database text-purple-500 text-xl"></i>
                    </div>
                </div>
                <div class="mt-4">
                    <div class="relative pt-1">
                        <div class="overflow-hidden h-2 text-xs flex rounded bg-purple-200">
                            <div style="width:$([math]::Round(($dashboardMetrics.StorageDetails.TotalStorageUsedGB / $dashboardMetrics.StorageDetails.TotalStorageAvailableGB) * 100))%"
                                class="shadow-none flex flex-col text-center whitespace-nowrap text-white justify-center bg-purple-500">
                            </div>
                        </div>
                    </div>
                </div>
            </div>

            <!-- License Card -->
            <div class="stat-card">
                <div class="flex items-center justify-between">
                    <div>
                        <p class="text-gray-500 text-sm font-medium">License Usage</p>
                        <p class="text-2xl font-bold text-gray-900">$($dashboardMetrics.TotalLicenses)</p>
                    </div>
                    <div class="rounded-full p-3 bg-green-100">
                        <i class="fas fa-key text-green-500 text-xl"></i>
                    </div>
                </div>
                <div class="mt-4">
                    <div class="flex justify-between text-sm">
                        <span class="text-green-500">Active: $($dashboardMetrics.TotalLicenses - $dashboardMetrics.UnusedLicenses)</span>
                        <span class="text-red-500">Unused: $($dashboardMetrics.UnusedLicenses)</span>
                    </div>
                </div>
            </div>
        </div>

        <!-- Storage Analysis -->
        <div class="bg-white rounded-lg shadow-sm p-6 mb-8">
            <h2 class="text-lg font-medium text-gray-900 mb-6">Storage Analysis</h2>
            
            <!-- Storage Type Distribution -->
            <div class="grid grid-cols-1 md:grid-cols-3 gap-4 mb-6">
                <div class="bg-gray-50 rounded-lg p-4">
                    <div class="flex justify-between items-center">
                        <span class="text-sm font-medium text-gray-500">Document Libraries</span>
                        <span class="text-lg font-semibold text-blue-600">
                            $($dashboardMetrics.StorageDetails.StorageTypes.DocumentLibraries) GB
                        </span>
                    </div>
                </div>
                <div class="bg-gray-50 rounded-lg p-4">
                    <div class="flex justify-between items-center">
                        <span class="text-sm font-medium text-gray-500">OneDrive</span>
                        <span class="text-lg font-semibold text-blue-600">
                            $($dashboardMetrics.StorageDetails.StorageTypes.OneDrive) GB
                        </span>
                    </div>
                </div>
                <div class="bg-gray-50 rounded-lg p-4">
                    <div class="flex justify-between items-center">
                        <span class="text-sm font-medium text-gray-500">SharePoint</span>
                        <span class="text-lg font-semibold text-blue-600">
                            $($dashboardMetrics.StorageDetails.StorageTypes.SharePoint) GB
                        </span>
                    </div>
                </div>
            </div>

            <!-- Storage Warnings -->
            $(if ($dashboardMetrics.StorageDetails.QuotaWarnings.Count -gt 0) {
@"
            <div class="bg-red-50 border-l-4 border-red-400 p-4 mb-6">
                <div class="flex">
                    <div class="flex-shrink-0">
                        <i class="fas fa-exclamation-triangle text-red-400"></i>
                    </div>
                    <div class="ml-3">
                        <h3 class="text-sm font-medium text-red-800">Storage Warnings</h3>
                        <div class="mt-2 text-sm text-red-700">
                            <ul class="list-disc pl-5 space-y-1">
                            $(foreach ($warning in $dashboardMetrics.StorageDetails.QuotaWarnings) {
                                "<li>$($warning.SiteName) is at $($warning.PercentageUsed)% capacity ($($warning.UsedGB)GB of $($warning.TotalGB)GB)</li>"
                            })
                            </ul>
                        </div>
                    </div>
                </div>
            </div>
"@
            })

            <!-- Storage Details Table -->
            <div class="overflow-x-auto">
                <table class="min-w-full divide-y divide-gray-200">
                    <thead class="bg-gray-50">
                        <tr>
                            <th scope="col" class="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Site</th>
                            <th scope="col" class="px-6 py-3 text-right text-xs font-medium text-gray-500 uppercase tracking-wider">Used (GB)</th>
                            <th scope="col" class="px-6 py-3 text-right text-xs font-medium text-gray-500 uppercase tracking-wider">Total (GB)</th>
                            <th scope="col" class="px-6 py-3 text-right text-xs font-medium text-gray-500 uppercase tracking-wider">Usage</th>
                            <th scope="col" class="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Last Modified</th>
                        </tr>
                    </thead>
                    <tbody class="bg-white divide-y divide-gray-200">
                        $(foreach ($site in $dashboardMetrics.StorageDetails.SiteDetails) {
                            $usageClass = if ($site.PercentageUsed -gt 90) { "text-red-600" } 
                                elseif ($site.PercentageUsed -gt 75) { "text-yellow-600" } 
                                else { "text-green-600" }
@"
                        <tr class="hover:bg-gray-50">
                            <td class="px-6 py-4 whitespace-nowrap text-sm font-medium text-gray-900">$($site.SiteName)</td>
                            <td class="px-6 py-4 whitespace-nowrap text-sm text-gray-500 text-right">$($site.UsedGB)</td>
                            <td class="px-6 py-4 whitespace-nowrap text-sm text-gray-500 text-right">$($site.TotalGB)</td>
                            <td class="px-6 py-4 whitespace-nowrap text-sm $usageClass text-right">$($site.PercentageUsed)%</td>
                            <td class="px-6 py-4 whitespace-nowrap text-sm text-gray-500">$($site.LastModified)</td>
                        </tr>
"@
                        })
                    </tbody>
                </table>
            </div>
        </div>

        <!-- Charts Section -->
        <div class="grid grid-cols-1 lg:grid-cols-2 gap-8 mb-8">
        <!-- Usage Distribution Chart -->
            <div class="bg-white rounded-lg shadow-sm p-6">
                <h3 class="text-lg font-medium text-gray-900 mb-4">Storage Distribution</h3>
                <div class="chart-container">
                    <canvas id="storageDistribution"></canvas>
                </div>
            </div>

            <!-- Storage Trends Chart -->
            <div class="bg-white rounded-lg shadow-sm p-6">
                <h3 class="text-lg font-medium text-gray-900 mb-4">License Distribution</h3>
                <div class="chart-container">
                    <canvas id="licenseDistribution"></canvas>
                </div>
            </div>
        </div>

        <!-- Security Section -->
        <div class="bg-white rounded-lg shadow-sm p-6 mb-8">
            <h2 class="text-lg font-medium text-gray-900 mb-6">Security Overview</h2>
            <div class="grid grid-cols-1 md:grid-cols-3 gap-4">
                <!-- MFA Status -->
                <div class="bg-gray-50 rounded-lg p-4">
                    <h3 class="text-sm font-medium text-gray-500">MFA Adoption</h3>
                    <p class="mt-2 text-3xl font-bold text-gray-900">
                        $([math]::Round(($dashboardMetrics.UsersWithMFA / $dashboardMetrics.TotalUsers) * 100))%
                    </p>
                    <p class="mt-1 text-sm text-gray-500">
                        $($dashboardMetrics.UsersWithMFA) of $($dashboardMetrics.TotalUsers) users
                    </p>
                </div>

                <!-- Device Compliance -->
                <div class="bg-gray-50 rounded-lg p-4">
                    <h3 class="text-sm font-medium text-gray-500">Device Compliance</h3>
                    <p class="mt-2 text-3xl font-bold text-gray-900">
                        $([math]::Round(($dashboardMetrics.CompliantDevices / $dashboardMetrics.TotalDevices) * 100))%
                    </p>
                    <p class="mt-1 text-sm text-gray-500">
                        $($dashboardMetrics.CompliantDevices) of $($dashboardMetrics.TotalDevices) devices
                    </p>
                </div>

                <!-- Security Score -->
                <div class="bg-gray-50 rounded-lg p-4">
                    <h3 class="text-sm font-medium text-gray-500">Security Score</h3>
                    <p class="mt-2 text-3xl font-bold text-gray-900">
                        $($dashboardMetrics.SecurityScore)%
                    </p>
                    <p class="mt-1 text-sm text-gray-500">
                        Overall security posture
                    </p>
                </div>
            </div>
        </div>
    </main>

    <!-- Footer -->
    <footer class="bg-white border-t border-gray-200">
        <div class="max-w-7xl mx-auto py-4 px-4 sm:px-6 lg:px-8">
            <p class="text-center text-sm text-gray-500">
                Generated using Microsoft Graph API on $($enhancedSummary.LastReportGeneration)
            </p>
        </div>
    </footer>

    <!-- JavaScript for Charts -->
    <script>
        // Initialize Charts
        document.addEventListener('DOMContentLoaded', function() {
            // Storage Distribution Chart
            const storageCtx = document.getElementById('storageDistribution').getContext('2d');
            new Chart(storageCtx, {
                type: 'doughnut',
                data: {
                    labels: ['Document Libraries', 'OneDrive', 'SharePoint'],
                    datasets: [{
                        data: [
                            $($dashboardMetrics.StorageDetails.StorageTypes.DocumentLibraries),
                            $($dashboardMetrics.StorageDetails.StorageTypes.OneDrive),
                            $($dashboardMetrics.StorageDetails.StorageTypes.SharePoint)
                        ],
                        backgroundColor: [
                            '#60A5FA',
                            '#34D399',
                            '#A78BFA'
                        ]
                    }]
                },
                options: {
                    responsive: true,
                    maintainAspectRatio: false,
                    plugins: {
                        legend: {
                            position: 'bottom'
                        }
                    }
                }
            });

            // License Distribution Chart
            const licenseCtx = document.getElementById('licenseDistribution').getContext('2d');
            new Chart(licenseCtx, {
                type: 'bar',
                data: {
                    labels: [$(foreach($license in $dashboardMetrics.LicenseDetails) { "'$($license.Name)'," })],
                    datasets: [{
                        label: 'Used Licenses',
                        data: [$(foreach($license in $dashboardMetrics.LicenseDetails) { "$($license.Used)," })],
                        backgroundColor: '#60A5FA'
                    },
                    {
                        label: 'Available Licenses',
                        data: [$(foreach($license in $dashboardMetrics.LicenseDetails) { "$($license.Total - $license.Used)," })],
                        backgroundColor: '#E5E7EB'
                    }]
                },
                options: {
                    responsive: true,
                    maintainAspectRatio: false,
                    scales: {
                        x: {
                            stacked: true
                        },
                        y: {
                            stacked: true
                        }
                    }
                }
            });
        });

        // Add responsive handlers
        window.addEventListener('resize', function() {
            const tables = document.querySelectorAll('table');
            tables.forEach(table => {
                const wrapper = table.parentElement;
                if(wrapper.scrollWidth > wrapper.clientWidth) {
                    wrapper.style.overflowX = 'auto';
                }
            });
        });

        // Print handler
        function beforePrint() {
            // Adjust chart sizes for printing
            Chart.instances.forEach(chart => {
                chart.resize();
            });
        }
        window.onbeforeprint = beforePrint;
    </script>
</body>
</html>
"@

# Export HTML report
$htmlReport | Out-File "Reports\Report.html"

Write-Host "`nReport Generation Complete!"
Write-Host "Reports are organized in the following structure:"
foreach ($category in $reportFolders.Keys) {
    Write-Host "`n$category"
    foreach ($subcategory in $reportFolders[$category]) {
        Write-Host "  ├─ $subcategory"
    }
}

Write-Host "`nHTML report is available at 'Reports\Report.html'"
Write-Host "Each CSV report includes a timestamp in the filename"