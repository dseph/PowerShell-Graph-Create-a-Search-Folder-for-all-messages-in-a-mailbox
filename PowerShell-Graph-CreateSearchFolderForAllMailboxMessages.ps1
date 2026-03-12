# PowerShell-Graph-CreateSearchFolderForAllMailboxMessages.ps1
#
# Generated initially using Copilot, then heavily edited by hand to be more complete and accurate. The original Copilot 
# output was a very basic skeleton that did not include the necessary details for creating a mailSearchFolder in 
# Microsoft Graph, such as the required filterQuery and the use of msgfolderroot for sourceFolderIds. I added detailed 
# comments, error handling, and an example of how to list messages from the created Search Folder to demonstrate that it works. 
# I also included references to the official Microsoft Graph documentation for further reading.     

# Cautions:  
# 1) There are costs: While search folders can be very useful, creating one with a very broad filter (like "all mail") can lead to 
# performance issues in real usage. Always tailor the filter to your needs and test with smaller scopes before going broad.
# After the search folder is created, it can take some time for the search folder to populate with results, especially if the 
# filter is broad. Only create a search folder if is absolutely necessary, and be mindful of the potential performance 
# implications - Exchange has to maintain search folders, so there is a cost to Exchange.  When you no longer need the 
# search folder, delete it to avoid unnecessary overhead.
# 2) It takes time: It takes time for a search folder to be populated initially.  Be patient and check back later to see the results 
# in Outlook or via Graph. It can take minutes to a day before the search folder is fully populated, depending on the 
# size of the mailbox and the filter used. If you try to list messages from the search folder immediately after creation, you 
# may get zero results until the indexing is complete. This is expected behavior. 
# 3) Be EXTREMELY CAREFUL with filtering. It is extremely easy to include far more types of items than intended.
# For example, your filter will include item types at the start and you will need to filter it down to email, a type of
# email, etc.  If you don't then the search folder will include contact, events and a lot more.  You can check
# the included content by looking at search folder in Outlook. Note though that Outlook may not show all that the filter covers because
# they are not syn'd to outlook due to delayed syncing or that its not showing the messages because they are very old - Outlook should show a link 
# per folder to click on to sync the data into Outlook from server.  One thing which looks like a message in the inbox is NDR report - it looks like a
# message but its not really - its IPM type is REPORT.IPM.Note.NDR - yes, it does not start with IPM.Note.  
# 4) Do a lot of research and testing before you decide to create and use a search folder in production.

# Basic information...
# This is a PowerShell script to create a mailSearchFolder in Microsoft Graph. The folder will be called ""All Mail (Graph Search Folder)" and will 
# be created under the "Search Folders" container in the specified user's mailbox. 
# The search filter will filter for for only messages (the item class starts with IPM.Note) and which are visible.
# The script also demonstrates how to list the top 10 messages from the created search folder 
#  It can also list search folders, delete the search filter and get the ID of the search filter by its name.
#  To use, configure the settings at the start of the script and then run and debug.  It displays a menu for choosing different actions.
# Prerequisites:
# - An Azure AD app with the Mail.ReadWrite application permission (app-only).
# - Don't forget to do an Admin Grant for the permissions.

# References:
# [1] Create mailSearchFolder: https://learn.microsoft.com/en-us/graph/api/mailsearchfolder-post?view=graph-rest-1.0
# [2] mailSearchFolder resource: https://learn.microsoft.com/en-us/graph/api/resources/mailsearchfolder?view=graph-rest-1.0     
# Note: This script is for demonstration purposes. Creating a Search Folder with a very broad filter (like "all mail") can lead to 
# performance issues in real usage. Always tailor the filter to your needs.

# =========================
# CONFIG (edit these)
# =========================
$TenantId     = "YOUR_TENANT_ID"
$ClientId     = "YOUR_APP_CLIENT_ID"
$ClientSecret = "YOUR_APP_CLIENT_SECRET"

# Target mailbox to create the search folder in:
$UserUPN      = "user@contoso.com"

# Name of the Search Folder as it will appear:
$SearchFolderDisplayName = "All Mail (Graph Search Folder)"

# IMPORTANT:
# - This filterQuery is REQUIRED. It’s an OData filter for messages. [1](https://learn.microsoft.com/en-us/graph/api/mailsearchfolder-post?view=graph-rest-1.0)

# Filter for only messages (the itemclass starts with IPM.Note) and which are visible.
$FilterQuery="singleValueExtendedProperties/Any(ep: ep/id eq 'String 0x001A' and startswith(ep/value,'IPM.Note'))&includeHiddenMessages=false"

# Where to CREATE the search folder:
# - Outlook for Windows expects it under WellKnownFolderName.SearchFolders. [2](https://learn.microsoft.com/en-us/graph/api/resources/mailsearchfolder?view=graph-rest-1.0)
# - In Graph, you can reference the well-known folder name "searchfolders" in the URL. [1](https://learn.microsoft.com/en-us/graph/api/mailsearchfolder-post?view=graph-rest-1.0)
$ParentFolderWellKnownName = "searchfolders"

# Graph base
$GraphBase = "https://graph.microsoft.com/v1.0"

# =========================
# Helper: Get app-only token
# =========================
function Get-GraphAppToken {
    param(
        [Parameter(Mandatory=$true)] [string] $TenantId,
        [Parameter(Mandatory=$true)] [string] $ClientId,
        [Parameter(Mandatory=$true)] [string] $ClientSecret
    )

    $tokenEndpoint = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"

    $body = @{
        client_id     = $ClientId
        client_secret = $ClientSecret
        scope         = "https://graph.microsoft.com/.default"
        grant_type    = "client_credentials"
    }

    $resp = Invoke-RestMethod -Method Post -Uri $tokenEndpoint -Body $body -ContentType "application/x-www-form-urlencoded"
    return $resp.access_token
}

# =========================
# Helper: Invoke Graph REST
# =========================
function Invoke-Graph {
    param(
        [Parameter(Mandatory=$true)] [ValidateSet("GET","POST","PATCH","DELETE")] [string] $Method,
        [Parameter(Mandatory=$true)] [string] $Uri,
        [Parameter(Mandatory=$true)] [string] $AccessToken,
        [object] $Body = $null
    )

    $headers = @{
        Authorization = "Bearer $AccessToken"
        "Content-Type" = "application/json"
    }

    if ($null -ne $Body) {
        $json = $Body | ConvertTo-Json -Depth 20
        return Invoke-RestMethod -Method $Method -Uri $Uri -Headers $headers -Body $json
    }
    else {
        return Invoke-RestMethod -Method $Method -Uri $Uri -Headers $headers
    }
}

# =============================
# Menu
# =============================
function ShowMainMenu{
    # colors https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.utility/write-host?view=powershell-7.5
    Write-Host ""
    Write-Host "========================================" -ForegroundColor yellow
    Write-Host "                Main Menu               " -ForegroundColor yellow
    Write-Host "========================================" -ForegroundColor yellow
    Write-Host "  1. Create Search Folder               " -ForegroundColor yellow  
    Write-Host "  2. Read 10 items from Search Folder   " -ForegroundColor yellow
    Write-Host "  3. List all Search Folders            " -ForegroundColor yellow
    Write-Host "  4. Get the ID of the Search Folder    " -ForegroundColor yellow
    Write-Host "  5  Delete Search Folder               " -ForegroundColor yellow
    Write-Host "  6. Exit                               " -ForegroundColor red  
    Write-Host "========================================" -ForegroundColor yellow

    # Prompt user for a choice
    $choice = Read-Host "Enter the number of your choice (1-6)"

    switch ($choice) {
        "1" {
            Write-Host ""
            Write-Host "Chose toCreate Search Folder" -ForegroundColor green
            Write-Host ""
            CreateSearchFolder
        }
        "2" {
            Write-Host ""
            Write-Host "Chose to Show 10 items from Search Folder" -ForegroundColor green
            Write-Host ""
            $FoundFolderId = GetIdOfSearchFolderByName -FolderName $SearchFolderDisplayName 
            if ($null -ne $FoundFolderId) {
                 
                #Write-Host "Deleting Folder: $FoundFolderId" -ForegroundColor red
                ShowTenFolderMessages -FolderId $FoundFolderId
            }
            else {
                Write-Host "Search Folder was not found." -ForegroundColor red
            }
             
        }
       "3" {
            Write-Host ""
            Write-Host "List all Search Folders " -ForegroundColor green
            Write-Host ""
            ListAllSearchFolders
        }
       "4" {
            Write-Host ""
            Write-Host "Chose to get the ID of the Search Folder" -ForegroundColor green
            Write-Host ""
            GetIdOfSearchFolderByName -FolderName $SearchFolderDisplayName
        }
        "5" {
            Write-Host ""
            Write-Host "Chose to Delete the Search Folder" -ForegroundColor green
            Write-Host ""
            $FoundFolderId = GetIdOfSearchFolderByName -FolderName $SearchFolderDisplayName
            if ($null -ne $FoundFolderId) {
                 
                Write-Host "Deleting Folder: $FoundFolderId" -ForegroundColor red
                DeleteFolderById -FolderId $FoundFolderId
            }
            else {
                Write-Host "Cannot delete Search Folder because it was not found." -ForegroundColor red
            }
        }
        "6" {
            Write-Host ""
            Write-Host "Chose to Exit" -ForegroundColor red  
            Write-Host ""
        }
        default {
            Write-Host "Invalid selection. Please run the script again and choose 1-6."  -ForegroundColor red -ForegroundColor white
        }
    }

    return $choice

}

function ShowMainMenuLoop {
    $choice  = ShowMainMenu
    while ($choice -ne "6") {
        $choice = ShowMainMenu
    }
}
 
function CreateSearchFolder {
    
    # Note: Only run the creation part once. After the search folder is created, do NOT run it again without changing the name, or you'll get a "name already exists" error. If you want to create multiple search folders, change the $SearchFolderDisplayName for each one.
 
    Write-Host "Creating Search Folder '$SearchFolderDisplayName' in $UserUPN's mailbox with filter '$FilterQuery'..." -ForegroundColor yellow
    # 1) Get the folder id for msgfolderroot (root of the mail folder tree)
    #    We'll use this as the single sourceFolderIds entry, and enable includeNestedFolders=true
    #    so the search traverses all child folders. [1](https://learn.microsoft.com/en-us/graph/api/mailsearchfolder-post?view=graph-rest-1.0)[2](https://learn.microsoft.com/en-us/graph/api/resources/mailsearchfolder?view=graph-rest-1.0)
    
  
    Write-Host "Resolving msgfolderroot id for $UserUPN ..."
    $msgRootUri = "$GraphBase/users/$UserUPN/mailFolders/msgfolderroot?`$select=id"
    $msgRoot = Invoke-Graph -Method GET -Uri $msgRootUri -AccessToken $token
    $msgRootId = $msgRoot.id

    Write-Host "msgfolderroot id: $msgRootId"

    # 2) Create the mailSearchFolder under the SearchFolders container so it appears in Outlook for Windows. [2](https://learn.microsoft.com/en-us/graph/api/resources/mailsearchfolder?view=graph-rest-1.0)
    #    Endpoint: POST /users/{id}/mailFolders/{id}/childFolders [1](https://learn.microsoft.com/en-us/graph/api/mailsearchfolder-post?view=graph-rest-1.0)
    
    #$filter = "singleValueExtendedProperties/Any(ep: ep/id eq '$MessageClassPropId' and ep/value eq '$IpmClass')"

    
    $createUri = "$GraphBase/users/$UserUPN/mailFolders/$ParentFolderWellKnownName/childFolders"

    $body = @{
        "@odata.type"         = "microsoft.graph.mailSearchFolder"  # required type [1](https://learn.microsoft.com/en-us/graph/api/mailsearchfolder-post?view=graph-rest-1.0)
        displayName           = $SearchFolderDisplayName
        includeNestedFolders  = $true                               # deep traversal [1](https://learn.microsoft.com/en-us/graph/api/mailsearchfolder-post?view=graph-rest-1.0)[2](https://learn.microsoft.com/en-us/graph/api/resources/mailsearchfolder?view=graph-rest-1.0)
        sourceFolderIds       = @($msgRootId)                       # "all folders" via msgfolderroot + deep traversal
        filterQuery           = $FilterQuery                        # required filter [1](https://learn.microsoft.com/en-us/graph/api/mailsearchfolder-post?view=graph-rest-1.0)
    }

    Write-Host "Creating Search Folder '$SearchFolderDisplayName' under '$ParentFolderWellKnownName'..."  
    $created = Invoke-Graph -Method POST -Uri $createUri -AccessToken $token -Body $body

    Write-Host "Created Search Folder:"
    Write-Host ("  id: {0}" -f $created.id)
    Write-Host ("  displayName: {0}" -f $created.displayName)
    Write-Host ("  includeNestedFolders: {0}" -f $created.includeNestedFolders)
    Write-Host ("  filterQuery: {0}" -f $created.filterQuery)
 
    return $created
}

function ShowTenFolderMessages {
    param(
        [Parameter(Mandatory=$true)] [string] $FolderId
    )
    Write-Host ""
    Write-Host "Listing first 10 messages from the Search Folder..." -ForegroundColor yellow
     
    #$listUri = "$GraphBase/users/$UserUPN/mailFolders/$($created.id)/messages?`$top=10&`$select=subject,receivedDateTime,from"
    $listUri = "$GraphBase/users/$UserUPN/mailFolders/$($FolderId)/messages?`$top=10&`$select=subject,receivedDateTime,from"
    $messages = Invoke-Graph -Method GET -Uri $listUri -AccessToken $token
    
    foreach ($message in $messages.value) {
        Write-Host ("receivedDateTime: '{0}' Subject: {1}" -f $message.receivedDateTime, $message.subject)
    }
  
    return
}   

function DeleteFolderById {
    param(
        [Parameter(Mandatory=$true)] [string] $FolderId
    )
    Write-Host ""
 
    # To delete the search folder when you're done, you can use the following code (make sure to set the correct search folder ID):
    Write-Host ""
    Write-Host "Deleting Folder: $FolderId" -ForegroundColor yellow
    #$deleteUri = "$GraphBase/users/$UserUPN/mailFolders/$($created.id)"
    $deleteUri = "$GraphBase/users/$UserUPN/mailFolders/$($FolderId)"

    Invoke-Graph -Method DELETE -Uri $deleteUri -AccessToken $token
}   

function ListAllSearchFolders {
    Write-Host ""
    Write-Host "Listing all Search Folders in the mailbox..."
    $listUri = "$GraphBase/users/$UserUPN/mailFolders/searchfolders/childFolders?`$select=id,displayName,totalItemCount"
    $folders = Invoke-Graph -Method GET -Uri $listUri -AccessToken $token
 
    foreach ($folder in $folders.value) { 
        Write-Host ("Name: '{0}'"  -f $folder.displayName)
        Write-Host ("    id: '{0}'"  -f $folder.id)
        Write-Host ("    totalItemCount: '{0}'"  -f $folder.totalItemCount)
        Write-Host ""
    }
    
    return
}

function GetIdOfSearchFolderByName {
    param(
        [Parameter(Mandatory=$true)] [string] $FolderName
    )
    Write-Host ""
    Write-Host "Getting ID of Search Folder with name '$FolderName'..." -ForegroundColor yellow
    $listUri = "$GraphBase/users/$UserUPN/mailFolders/searchfolders/childFolders?`$select=id,displayName"
    $folders = Invoke-Graph -Method GET -Uri $listUri -AccessToken $token
    foreach ($folder in $folders.value) {
        if ($folder.displayName -eq $FolderName) {
            Write-Host ("Found folder '{0}' with id: {1}" -f $folder.displayName, $folder.id)
            return $folder.id
        }
    }
    Write-Host ("No Search Folder found with name '{0}'" -f $FolderName)
    return $null
}   
 
# =========================
# MAIN
# =========================

Write-Host "Getting app-only token..."
$token = Get-GraphAppToken -TenantId $TenantId -ClientId $ClientId -ClientSecret $ClientSecret
Write-Host "Token acquired." -ForegroundColor green
Write-Host ""
ShowMainMenuLoop
    

 
