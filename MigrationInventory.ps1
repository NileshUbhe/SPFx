<#
.SYNOPSIS
    SharePoint Online Site Inventory Script with Certificate Authentication
    Pfx file               : c:\temp\PnPAutomationScript.pfx
    Cer file               : c:\temp\PnPAutomationScript.cer
    AzureAppId/ClientId    : 878f555b-acb4-47de-a3d0-ba05faec366a
    Certificate Thumbprint : 2F1A99BAD4F64A6840F476FBEE13B3E22EB14CD4
.DESCRIPTION
    Iterates through all SharePoint Online sites and subsites recursively to collect:
    - Site and subsite details
    - Site templates
    - Lists and libraries with item counts
    - Permission matrix including unique permissions
.NOTES
    Requires PnP.PowerShell module and CommonFunctions.ps1
#>

# Import common functions
. "$PSScriptRoot\CommonFunctions.ps1"

#region Configuration
$config = @{
    TenantAdminUrl      = "https://0nthj-admin.sharepoint.com"
    ClientId            = "878f555b-acb4-47de-a3d0-ba05faec366a"
    CertificatePath     = "c:\temp\PnPAutomationScript.pfx"
    CertificatePassword = "S#curityp03"
    Tenant              = "0nthj.onmicrosoft.com"
    OutputPath          = "$PSScriptRoot\InventoryReports"
    LogPath             = "$PSScriptRoot\Logs"
}
#endregion

#region Helper Functions

function Initialize-InventoryReporting {
    param (
        [string]$OutputPath,
        [string]$LogPath
    )

    # Create output directories if they don't exist
    if (-not (Test-Path $OutputPath)) {
        New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
    }

    if (-not (Test-Path $LogPath)) {
        New-Item -ItemType Directory -Path $LogPath -Force | Out-Null
    }

    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    
    return @{
        SitesReport       = "$OutputPath\Sites_$timestamp.csv"
        ListsReport       = "$OutputPath\Lists_$timestamp.csv"
        PermissionsReport = "$OutputPath\Permissions_$timestamp.csv"
        ErrorLog          = "$LogPath\Errors_$timestamp.log"
        ExecutionLog      = "$LogPath\Execution_$timestamp.log"
    }
}

function Write-Log {
    param (
        [string]$Message,
        [ValidateSet('Info', 'Warning', 'Error')]
        [string]$Level = 'Info',
        [string]$LogPath
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"
    
    Add-Content -Path $LogPath -Value $logMessage
    
    switch ($Level) {
        'Info' { Write-Host $logMessage -ForegroundColor Green }
        'Warning' { Write-Host $logMessage -ForegroundColor Yellow }
        'Error' { Write-Host $logMessage -ForegroundColor Red }
    }
}

function Get-SPOSiteInventory {
    param (
        [Parameter(Mandatory = $true)]
        [string]$SiteUrl,
        [Parameter(Mandatory = $true)]
        [hashtable]$Config,
        [Parameter(Mandatory = $true)]
        [hashtable]$Reports
    )
    
    try {
        Write-Log -Message "Processing site: $SiteUrl" -Level Info -LogPath $Reports.ExecutionLog
        
        # Connect to site
        $connection = Connect-SP-Online-WithRetry -DestinationUrl $SiteUrl `
            -ClientId $Config.ClientId `
            -CertificatePath $Config.CertificatePath `
            -CertificatePwd $Config.CertificatePassword `
            -Tenant $Config.Tenant
        
        # Get site details
        $site = Get-PnPSite -Connection $connection -Includes RootWeb, Owner
        $web = Get-PnPWeb -Connection $connection -Includes WebTemplate, ServerRelativeUrl, HasUniqueRoleAssignments
        
        # Collect site information
        $siteInfo = [PSCustomObject]@{
            SiteUrl              = $SiteUrl
            Title                = $web.Title
            Template             = $web.WebTemplate
            TemplateId           = "$($web.WebTemplate)#$($web.Configuration)"
            ServerRelativeUrl    = $web.ServerRelativeUrl
            Created              = $web.Created
            LastItemModifiedDate = $web.LastItemModifiedDate
            HasUniquePermissions = $web.HasUniqueRoleAssignments
            WebId                = $web.Id
            SiteId               = $site.Id
        }
        
        # Export site info
        $siteInfo | Export-Csv -Path $Reports.SitesReport -Append -NoTypeInformation
        
        # Get permissions for the site
        $siteTitle = if ([string]::IsNullOrWhiteSpace($web.Title)) { $SiteUrl } else { $web.Title }
        Get-SPOPermissions -WebUrl $SiteUrl -Connection $connection -Reports $Reports -ObjectType "Site" -ObjectTitle $siteTitle
        

        # Get lists and libraries
        Get-SPOListsInventory -Connection $connection -WebUrl $SiteUrl -Reports $Reports
        
        $customReport = $Reports.SitesReport -replace 'Sites_', 'CustomSolutions_'

        # User custom actions (site and web scope)
        foreach ($scope in @('Site', 'Web')) {
            $customActions = Get-PnPCustomAction -Connection $connection -Scope $scope -ErrorAction SilentlyContinue
            foreach ($action in $customActions) {
            [PSCustomObject]@{
                SiteUrl      = $SiteUrl
                Type         = 'UserCustomAction'
                Scope        = $scope
                Title        = $action.Title
                Name         = $action.Name
                Location     = $action.Location
                ScriptSrc    = $action.ScriptSrc
                ScriptBlock  = $action.ScriptBlock
                Description  = $action.Description
                Id           = $action.Id
            } | Export-Csv -Path $customReport -Append -NoTypeInformation -Force
            }
        }

        # Installed apps/add-ins
        $apps = Get-PnPApp -Connection $connection -Scope Site -ErrorAction SilentlyContinue |
            Where-Object { $_.InstalledVersion }
        foreach ($app in $apps) {
            [PSCustomObject]@{
            SiteUrl          = $SiteUrl
            Type             = 'App'
            Scope            = 'Site'
            Title            = $app.Title
            Name             = $app.Title
            Id               = $app.Id
            ProductId        = $app.ProductId
            Publisher        = $app.Publisher
            InstalledVersion = $app.InstalledVersion
            Source           = $app.Source
            } | Export-Csv -Path $customReport -Append -NoTypeInformation
        }

        # Client-side web parts on modern pages
        $sitePages = Get-PnPListItem -List "Site Pages" -Connection $connection -Fields "FileRef" -ErrorAction SilentlyContinue
        # Include items from Pages library as well
        $pagesLibrary = Get-PnPListItem -List "Pages" -Connection $connection -Fields "FileRef" -ErrorAction SilentlyContinue
        
        if ($pagesLibrary) {
            $sitePages += $pagesLibrary
        }
        foreach ($pageItem in $sitePages) {
            $page = Get-PnPClientSidePage -Identity $pageItem.FieldValues.FileRef -Connection $connection -ErrorAction SilentlyContinue
            if ($page -eq $null) {
                continue
            }
            foreach ($control in $page.Controls | Where-Object { $_.GetType().Name -eq 'ClientSideWebPart' }) {
            [PSCustomObject]@{
                SiteUrl     = $SiteUrl
                Type        = 'ClientSideWebPart'
                PageUrl     = $pageItem.FieldValues.FileRef
                Title       = $control.Title
                Name        = $control.Title
                WebPartId   = $control.WebPartId
                InstanceId  = $control.InstanceId
                ComponentId = $control.ComponentId
            } | Export-Csv -Path $customReport -Append -NoTypeInformation
            }
        }
        # Get subsites recursively
        $subWebs = Get-PnPSubWeb -Connection $connection -Recurse
        
        if ($subWebs) {
            Write-Log -Message "Found $($subWebs.Count) subsite(s) under $SiteUrl" -Level Info -LogPath $Reports.ExecutionLog
            
            foreach ($subWeb in $subWebs) {
                $subSiteUrl = $subWeb.Url
                Get-SPOSiteInventory -SiteUrl $subSiteUrl -Config $Config -Reports $Reports
            }
        }
        
        #Disconnect-PnPOnline -Connection $connection
        
    }
    catch {
        $errorMessage = "Error processing site $SiteUrl : $($_.Exception.Message)"
        Write-Log -Message $errorMessage -Level Error -LogPath $Reports.ErrorLog
    }
}

function Get-SPOListsInventory {
    param (
        [Parameter(Mandatory = $true)]
        $Connection,
        [Parameter(Mandatory = $true)]
        [string]$WebUrl,
        [Parameter(Mandatory = $true)]
        [hashtable]$Reports
    )
    
    try {
        $lists = Get-PnPList -Connection $Connection -Includes ItemCount, HasUniqueRoleAssignments, RootFolder, BaseTemplate
        
        foreach ($list in $lists) {
            # Skip system lists if needed
            if ($list.Hidden -eq $false) {
                
                $listInfo = [PSCustomObject]@{
                    SiteUrl              = $WebUrl
                    ListTitle            = $list.Title
                    ListUrl              = $list.RootFolder.ServerRelativeUrl
                    BaseTemplate         = $list.BaseTemplate
                    ItemCount            = $list.ItemCount
                    HasUniquePermissions = $list.HasUniqueRoleAssignments
                    Created              = $list.Created
                    LastItemModifiedDate = $list.LastItemModifiedDate
                    EnableVersioning     = $list.EnableVersioning
                    ListId               = $list.Id
                }
                
                $listInfo | Export-Csv -Path $Reports.ListsReport -Append -NoTypeInformation
                
                # Get list permissions
                Get-SPOPermissions -WebUrl $WebUrl -Connection $Connection -Reports $Reports `
                    -ObjectType "List" -ObjectTitle $list.Title -ListTitle $list.Title
                
                # Check for items with unique permissions
                if ($list.ItemCount -gt 0 -and $list.ItemCount -lt 5000) {
                    Get-SPOItemUniquePermissions -Connection $Connection -List $list -WebUrl $WebUrl -Reports $Reports
                }
                elseif ($list.ItemCount -ge 5000) {
                    Write-Log -Message "List '$($list.Title)' has $($list.ItemCount) items. Skipping item-level permission check." `
                        -Level Warning -LogPath $Reports.ExecutionLog
                }
            }
        }
    }
    catch {
        $errorMessage = "Error processing lists for $WebUrl : $($_.Exception.Message)"
        Write-Log -Message $errorMessage -Level Error -LogPath $Reports.ErrorLog
    }
}

function Get-SPOItemUniquePermissions {
    param (
        [Parameter(Mandatory = $true)]
        $Connection,
        [Parameter(Mandatory = $true)]
        $List,
        [Parameter(Mandatory = $true)]
        [string]$WebUrl,
        [Parameter(Mandatory = $true)]
        [hashtable]$Reports
    )
    
    try {
        $items = Get-PnPListItem -List $List -Connection $Connection -Fields "ID", "FileLeafRef", "FileRef", "HasUniqueRoleAssignments"
        
        foreach ($item in $items) {
            if ($item.FieldValues.HasUniqueRoleAssignments -eq $true) {
                $itemTitle = if ($item.FieldValues.FileLeafRef) { $item.FieldValues.FileLeafRef } else { "Item ID: $($item.Id)" }
                
                Get-SPOPermissions -WebUrl $WebUrl -Connection $Connection -Reports $Reports `
                    -ObjectType "ListItem" -ObjectTitle $itemTitle -ListTitle $List.Title -ItemId $item.Id
            }
        }
    }
    catch {
        $errorMessage = "Error checking item permissions for list '$($List.Title)': $($_.Exception.Message)"
        Write-Log -Message $errorMessage -Level Error -LogPath $Reports.ErrorLog
    }
}

function Get-SPOPermissions {
    param (
        [Parameter(Mandatory = $true)]
        [string]$WebUrl,
        [Parameter(Mandatory = $true)]
        $Connection,
        [Parameter(Mandatory = $true)]
        [hashtable]$Reports,
        [Parameter(Mandatory = $true)]
        [ValidateSet('Site', 'List', 'ListItem')]
        [string]$ObjectType,
        [Parameter(Mandatory = $true)]
        [string]$ObjectTitle,
        [string]$ListTitle = "",
        [int]$ItemId = 0
    )
    
    try {
        $roleAssignments = $null
        
        switch ($ObjectType) {
            'Site' {
                $roleAssignments = Get-PnPProperty -ClientObject (Get-PnPWeb -Connection $Connection) -Property RoleAssignments -Connection $Connection
            }
            'List' {
                $list = Get-PnPList -Identity $ListTitle -Connection $Connection
                $roleAssignments = Get-PnPProperty -ClientObject $list -Property RoleAssignments -Connection $Connection
            }
            'ListItem' {
                $list = Get-PnPList -Identity $ListTitle -Connection $Connection
                $item = Get-PnPListItem -List $list -Id $ItemId -Connection $Connection
                $roleAssignments = Get-PnPProperty -ClientObject $item -Property RoleAssignments -Connection $Connection
            }
        }
        
        if ($roleAssignments) {
            foreach ($roleAssignment in $roleAssignments) {
                $member = Get-PnPProperty -ClientObject $roleAssignment -Property Member -Connection $Connection
                $roleDefinitions = Get-PnPProperty -ClientObject $roleAssignment -Property RoleDefinitionBindings -Connection $Connection
                
                foreach ($roleDef in $roleDefinitions) {
                    $permissionInfo = [PSCustomObject]@{
                        SiteUrl            = $WebUrl
                        ObjectType         = $ObjectType
                        ObjectTitle        = $ObjectTitle
                        ListTitle          = $ListTitle
                        ItemId             = if ($ItemId -gt 0) { $ItemId } else { "" }
                        PrincipalType      = $member.PrincipalType
                        PrincipalName      = $member.Title
                        PrincipalLoginName = $member.LoginName
                        PermissionLevel    = $roleDef.Name
                        PermissionType     = $roleDef.BasePermissions
                    }
                    
                    $permissionInfo | Export-Csv -Path $Reports.PermissionsReport -Append -NoTypeInformation
                }
            }
        }
    }
    catch {
        $errorMessage = "Error getting permissions for $ObjectType '$ObjectTitle': $($_.Exception.Message)"
        Write-Log -Message $errorMessage -Level Error -LogPath $Reports.ErrorLog
    }
}

#endregion

#region Main Execution

function Start-SPOInventory {
    param (
        [Parameter(Mandatory = $true)]
        [hashtable]$Config
    )
    
    Write-Host "=== SharePoint Online Site Inventory Script ===" -ForegroundColor Cyan
    Write-Host "Starting inventory process..." -ForegroundColor Cyan
    
    # Initialize reporting
    $reports = Initialize-InventoryReporting -OutputPath $Config.OutputPath -LogPath $Config.LogPath
    
    Write-Log -Message "Inventory process started" -Level Info -LogPath $reports.ExecutionLog
    
    try {
        # Connect to tenant admin
        Write-Log -Message "Connecting to SharePoint Admin Center..." -Level Info -LogPath $reports.ExecutionLog
        
        $adminConnection = Connect-SP-Online-WithRetry -DestinationUrl $Config.TenantAdminUrl `
            -ClientId $Config.ClientId `
            -CertificatePath $Config.CertificatePath `
            -CertificatePwd $Config.CertificatePassword `
            -Tenant $Config.Tenant
        
        # Get all site collections
        $sites = Get-PnPTenantSite -Connection $adminConnection | Where-Object { $_.Template -notlike "*Redirect*" }
        
        Write-Log -Message "Found $($sites.Count) site collection(s)" -Level Info -LogPath $reports.ExecutionLog
        
        #Disconnect-PnPOnline -Connection $adminConnection
        
        # Process each site collection
        $siteCounter = 0
        foreach ($site in $sites) {
            $siteCounter++
            Write-Host "`nProcessing site $siteCounter of $($sites.Count): $($site.Url)" -ForegroundColor Cyan
            
            Get-SPOSiteInventory -SiteUrl $site.Url -Config $Config -Reports $reports
        }
        
        Write-Log -Message "Inventory process completed successfully" -Level Info -LogPath $reports.ExecutionLog
        Write-Host "`n=== Inventory Complete ===" -ForegroundColor Green
        Write-Host "Reports generated at:" -ForegroundColor Green
        Write-Host "  Sites: $($reports.SitesReport)" -ForegroundColor White
        Write-Host "  Lists: $($reports.ListsReport)" -ForegroundColor White
        Write-Host "  Permissions: $($reports.PermissionsReport)" -ForegroundColor White
        Write-Host "  Logs: $($reports.ExecutionLog)" -ForegroundColor White
    }
    catch {
        $errorMessage = "Critical error in inventory process: $($_.Exception.Message)"
        Write-Log -Message $errorMessage -Level Error -LogPath $reports.ErrorLog
        throw
    }
}

# Execute the inventory
Start-SPOInventory -Config $config

#endregion