<#
.SYNOPSIS
    Updates the column visibility within the SharePoint Views/Forms
.DESCRIPTION
    This script can be used to update the visiblity of a columns within the SharePoint Views and the Forms.
.AUTHOR
    Rick Wilson
.PARAMETER SPSiteUri
    Enter the SharePoint site uri.
.PARAMETER ListName
    Enter the SharePoint List Name.
.PARAMETER FieldName
    Enter the SharePoint Field Name.
.PARAMETER ShowInEditForm
    Enter $true or $false to set the visiblity of this columns wihin an edit form.
.PARAMETER ShowInNewForm
    Enter $true or $false to set the visiblity of this columns wihin a new form.
.PARAMETER ShowInDisplayForm
    Enter $true or $false to set the visiblity of this columns wihin a display form.
.PARAMETER ShowInViewForms
    Enter $true or $false to set the visiblity of this columns wihin a list view.
.EXAMPLE
    PS> ./UpdateColumnVisibility.ps1 -SPSiteUri "https://yoursharepoint.sharepoint.com/sites/Clients" -ListName "Tickets" -FieldName "Status" -ShowInEditForm $false -ShowInNewForm $false -ShowInDisplayForm $true -ShowInViewForms $true
#>

Param(
    [Parameter(Mandatory=$true,
    ValueFromPipeline=$true)]
    [string]
    $SPSiteUri,

    [Parameter(Mandatory=$true,
    ValueFromPipeline=$true)]
    [string]
    $ListName,

    [Parameter(Mandatory=$true,
    ValueFromPipeline=$true)]
    [string]
    $FieldName,

    [Parameter(Mandatory=$true,
    ValueFromPipeline=$true)]
    [bool]
    $ShowInEditForm,

    [Parameter(Mandatory=$true,
    ValueFromPipeline=$true)]
    [bool]
    $ShowInNewForm,

    [Parameter(Mandatory=$true,
    ValueFromPipeline=$true)]
    [bool]
    $ShowInDisplayForm,

    [Parameter(Mandatory=$true,
    ValueFromPipeline=$true)]
    [bool]
    $ShowInViewForms
)

# Add the Microsoft SharePoint PowerShell components in case this script is not running withing the SharePoint Management console.
Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue

# Get the Web
$Web = Get-SPWeb $SPSiteUri

# Get the List
$List = $Web.Lists[$ListName]

# Get the Field 
$Field = $List.Fields[$FieldName]

# Update the Field display properties
$Field.ShowInEditForm = $ShowInEditForm
$Field.ShowInNewForm = $ShowInNewForm
$Field.ShowInDisplayForm = $ShowInDisplayForm
$Field.ShowInViewForms = $ShowInViewForms

#Complete updates and dispose
$Field.Update()
$Web.Update()
$Web.Dispose()