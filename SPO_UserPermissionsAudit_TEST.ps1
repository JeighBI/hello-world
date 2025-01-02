# Set parameter values
$SiteURL = "https://tenant-admin.sharepoint.com"
$Tenant = "tenant.onmicrosoft.com"
$ClientId = "c48f027c-9c6b-479d-a339-59b6a4f9bad8"
$CertPath = "C:\tmp\certificate.pfx"
$CertPassword = ConvertTo-SecureString -String "YourStrongPassword" -Force -AsPlainText
$UserAccount = "i:0#.f|membership|alan.turing@tenant.onmicrosoft.com"
$ReportFile = "C:\Users\jbaid\Downloads\SitePermissionRpt_010225.csv"
$BatchSize = 500

# Connect to SharePoint Online using modern authentication
Connect-PnPOnline -Tenant $Tenant -Url $SiteURL -ClientId $ClientId -CertificatePath $CertPath -CertificatePassword $CertPassword

# Example command to get all site collections
$SiteCollections = Get-PnPTenantSite


# Initialize an array to store the report data
$ReportData = New-Object System.Collections.Generic.List[PSObject]

# Function to get user permissions applied on a particular object, such as Web, List, Folder, or Item
Function Get-Permissions([Microsoft.SharePoint.Client.SecurableObject]$Object) {
    # Determine the type of the object
    Switch ($Object.TypedObject.ToString()) {
        "Microsoft.SharePoint.Client.Web" {
            $ObjectType = "Site"
            $ObjectURL = $Object.Url
            $ObjectTitle = $Object.Title
        }
        "Microsoft.SharePoint.Client.ListItem" {
            $ObjectType = "List Item/Folder"
            # Get the URL of the List Item
            $Object.ParentList.Retrieve("DefaultDisplayFormUrl")
            $Ctx.ExecuteQuery()
            $DefaultDisplayFormUrl = $Object.ParentList.DefaultDisplayFormUrl
            $ObjectURL = $("{0}{1}?ID={2}" -f $Ctx.Web.Url.Replace($Ctx.Web.ServerRelativeUrl, ''), $DefaultDisplayFormUrl, $Object.Id)
            # Retrieve the file name using FileLeafRef
            if ($Object.FieldValues.ContainsKey("FileLeafRef")) {
                $ObjectTitle = $Object["FileLeafRef"]
            } else {
                $ObjectTitle = "No Name"
            }
        }
        Default {
            $ObjectType = "List/Library"
            # Get the URL of the List or Library
            $Ctx.Load($Object.RootFolder)
            $Ctx.ExecuteQuery()
            $ObjectURL = $("{0}{1}" -f $Ctx.Web.Url.Replace($Ctx.Web.ServerRelativeUrl, ''), $Object.RootFolder.ServerRelativeUrl)
            $ObjectTitle = $Object.Title
        }
    }

    # Get permissions assigned to the object
    $Ctx.Load($Object.RoleAssignments)
    $Ctx.ExecuteQuery()

    Foreach ($RoleAssignment in $Object.RoleAssignments) {
        $Ctx.Load($RoleAssignment.Member)
        $Ctx.ExecuteQuery()

        # Check direct permissions
        if ($RoleAssignment.Member.PrincipalType -eq "User") {
            # Is the current user the user we search for?
            if ($RoleAssignment.Member.LoginName -eq $SearchUser.LoginName) {
                Write-Host -f Cyan "Found the User under direct permissions of the $($ObjectType) at $($ObjectURL)"

                # Get the Permissions assigned to user
                $UserPermissions = @()
                $Ctx.Load($RoleAssignment.RoleDefinitionBindings)
                $Ctx.ExecuteQuery()
                foreach ($RoleDefinition in $RoleAssignment.RoleDefinitionBindings) {
                    $UserPermissions += $RoleDefinition.Name + ";"
                }
                # Send the Data to Report file
                $ReportData.Add([PSCustomObject]@{
                    URL             = $ObjectURL
                    Object          = $ObjectType
                    Title           = $ObjectTitle
                    PermissionType  = "Direct Permission"
                    Permissions     = $UserPermissions -join ","
                })
            }
        }
        ElseIf ($RoleAssignment.Member.PrincipalType -eq "SharePointGroup") {
            # Search inside SharePoint Groups and check if the user is a member of that group
            $Group = $Web.SiteGroups.GetByName($RoleAssignment.Member.LoginName)
            $GroupUsers = $Group.Users
            $Ctx.Load($GroupUsers)
            $Ctx.ExecuteQuery()

            # Check if the user is a member of the group
            Foreach ($User in $GroupUsers) {
                # Check if the search user is a member of the group
                if ($User.LoginName -eq $SearchUser.LoginName) {
                    Write-Host -f Cyan "Found the User under Member of the Group '$($RoleAssignment.Member.LoginName)' on $($ObjectType) at $($ObjectURL)"

                    # Get the Group's Permissions on the site
                    $GroupPermissions = @()
                    $Ctx.Load($RoleAssignment.RoleDefinitionBindings)
                    $Ctx.ExecuteQuery()
                    Foreach ($RoleDefinition in $RoleAssignment.RoleDefinitionBindings) {
                        $GroupPermissions += $RoleDefinition.Name + ";"
                    }
                    # Send the Data to Report file
                    $ReportData.Add([PSCustomObject]@{
                        URL             = $ObjectURL
                        Object          = $ObjectType
                        Title           = $ObjectTitle
                        PermissionType  = "Member of '$($RoleAssignment.Member.LoginName)' Group"
                        Permissions     = $GroupPermissions -join ","
                    })
                }
            }
        }
    }
}

# Function to check permissions of all list items of a given list
Function Check-SPOListItemsPermission([Microsoft.SharePoint.Client.List]$List) {
    Write-Host -f Yellow "Searching in List Items of the List '$($List.Title)..."

    $Query = New-Object Microsoft.SharePoint.Client.CamlQuery
    $Query.ViewXml = "<View Scope='RecursiveAll'><Query><OrderBy><FieldRef Name='ID' Ascending='TRUE'/></OrderBy></Query><RowLimit Paged='TRUE'>$BatchSize</RowLimit></View>"

    $Counter = 0
    # Batch process list items - to mitigate list threshold issue on larger lists
    Do {
        # Get items from the list in Batch
        $ListItems = $List.GetItems($Query)
        $Ctx.Load($ListItems)
        $Ctx.ExecuteQuery()

        $Query.ListItemCollectionPosition = $ListItems.ListItemCollectionPosition
        # Loop through each List item
        ForEach ($ListItem in $ListItems) {
            $ListItem.Retrieve("HasUniqueRoleAssignments")
            $Ctx.ExecuteQuery()
            # Retrieve the file name using FileLeafRef
            if ($ListItem.FieldValues.ContainsKey("FileLeafRef")) {
                $ListItemTitle = $ListItem["FileLeafRef"]
            } else {
                $ListItemTitle = "No Name"
            }
            if ($ListItem.HasUniqueRoleAssignments -eq $true) {
                # Call the function to generate Permission report
                Get-Permissions -Object $ListItem
            }
            $Counter++
            Write-Progress -PercentComplete ($Counter / ($List.ItemCount) * 100) -Activity "Processing Items $Counter of $($List.ItemCount)" -Status "Searching Unique Permissions in List Items of $($List.Title)"
        }
    } While ($Query.ListItemCollectionPosition -ne $null)
}

# Function to check permissions of all lists from the web
Function Check-SPOListPermission([Microsoft.SharePoint.Client.Web]$Web) {
    # Get All Lists from the web
    $Lists = Get-PnPList
            #Exclude system lists
            $ExcludedLists = @("Access Requests","App Packages","appdata","appfiles","Apps in Testing","Cache Profiles","Composed Looks","Content and Structure Reports","Content type publishing error log","Device Channels","Form Templates","fpdatasources","Get started with Apps for Office and SharePoint","List Template Gallery", "Long Running Operation Status","Maintenance Log Library", "IMEDICT", "Master Docs","Master Page Gallery","MicroFeed","NintexFormXml","Quick Deploy Items","Relationships List","Reusable Content","Reporting Metadata", "Reporting Templates", "Search Config List", "Search Query Logs","Settings","Site Assets","Site Collection Documents","Site Collection Images","Site Pages", "Solution Gallery","Style Library","Suggested Content Browser Locations","Theme Gallery", "TaxonomyHiddenList","User Information List","Web Part Gallery","wfpub","wfsvc","Workflow History","Workflow Tasks")

    # Get all lists from the web
    ForEach ($List in $Lists) {
        # Exclude System Lists
        If ($List.BaseTemplate -eq 101 -and $List.Hidden -eq $False -and $ExcludedLists -notcontains $List.Title) {
            # Get List Items Permissions
            Check-SPOListItemsPermission $List

            # Get the Lists with Unique permission
            $List.Retrieve("HasUniqueRoleAssignments")
            $Ctx.ExecuteQuery()

            If ($List.HasUniqueRoleAssignments -eq $True) {
                # Call the function to check permissions
                Get-Permissions -Object $List
            }
        }
    }
}

# Function to check web's permissions from the given URL
Function Check-SPOWebPermission([Microsoft.SharePoint.Client.Web]$Web) {
    # Get all immediate subsites of the site
    $Webs = $Web.Webs
    $Ctx.Load($Webs)
    $Ctx.ExecuteQuery()

    # Call the function to get lists of the web
    Write-Host -f Yellow "Searching in the Web $($Web.Url)..."

    # Check if the Web has unique permissions
    $Web.Retrieve("HasUniqueRoleAssignments")
    $Ctx.ExecuteQuery()

    # Get the Web's Permissions
    If ($Web.HasUniqueRoleAssignments -eq $true) {
        Get-Permissions -Object $Web
    }

    # Scan Lists with Unique Permissions
    Write-Host -f Yellow "Searching in the Lists and Libraries of $($Web.Url)..."
    Check-SPOListPermission($Web)

    # Iterate through each subsite in the current web
    ForEach ($Subweb in $Webs) {
        # Call the function recursively
        Check-SPOWebPermission $Subweb
    }
}

# Iterate through each site collection
ForEach ($Site in $SiteCollections) {
    Write-Host "Processing site collection: $($Site.Url)"
    Connect-PnPOnline -Tenant $Tenant -Url $Site.Url -ClientId c48f027c-9c6b-479d-a339-59b6a4f9bad8 -CertificatePath $CertPath -CertificatePassword $CertPassword

    Try {
        # Setup the context
        $Ctx = Get-PnPContext

        # Get the Web
        $Web = Get-PnPWeb

        # Get the User object
        $SearchUser = Get-PnPUser -Identity $UserAccount

        # Call the function with RootWeb to get site collection permissions
        Check-SPOWebPermission $Web
    }
    Catch {
        Write-Host -f Red "Error processing site collection $($Site.Url): $_.Exception.Message"
    }
}

# Export the report data to a CSV file
$ReportData | Export-Csv -Path $ReportFile -NoTypeInformation

Write-Host -f Green "User Permission Report Generated Successfully!"