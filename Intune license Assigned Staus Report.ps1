<# 
.NOTES
Objective:       Script to used to find the Intune licence assignment status
Version:         1.0
Author:          Chander Mani Pandey
Creation Date:   16 March 2024
Find Author on 
Youtube:-        https://www.youtube.com/@chandermanipandey8763
Twitter:-        https://twitter.com/Mani_CMPandey
LinkedIn:-       https://www.linkedin.com/in/chandermanipandey

#>

#Note: this scrit is installing all Microsoft graph module but you can limit this to specific like Microsoft.Graph.Intune ,Microsoft.Graph.Authentication
cls
Set-ExecutionPolicy -ExecutionPolicy Bypass -Force

#-------------------------------- User Input Section Start---------------------------------------------------------------------#

$ReportingPath = "C:\TEMP\Licence_Report"            #Report Location
[string]$Platform = 'All'                            # Possible string :-> All,Windows,macos,iOS/iPadOS,Android (Personally-Owned Work Profile),Android (Device Administrator)

$ProductName1 = "Enterprise Mobility + Security E5"  # Provide Product names and service plan identifiers for licensing and update line number 78
$SKUID1 = "c1ec4a95-1f05-45b3-a911-aa3fa01094f5"     # Provide Product names and service plan identifiers for licensing and update line number 78

$ProductName2 = "Intune Plan 2"                      # Provide Product names and service plan identifiers for licensing and update line number 78
$SKUID2 = "d9923fe3-a2de-4d29-a5be-e3e83bb786be"     # Provide Product names and service plan identifiers for licensing and update line number 78


#-------------------------------- User Input Section End-----------------------------------------------------------------------#

# Check if the Microsoft.Graph module is installed
if (-not (Get-Module -Name Microsoft.Graph -ListAvailable)) {
    Write-Host "Microsoft.Graph module not found. Installing..."
    # Module is not installed, so install it
    Install-Module Microsoft.Graph -Scope CurrentUser -Force -AllowClobber
    #Install-Module Microsoft.Graph -Scope AllUsers -Force -AllowClobber
    Write-Host "Microsoft.Graph module installed successfully." -ForegroundColor Green
} else {
    Write-Host "Microsoft.Graph module is already installed." -ForegroundColor Green
}

Write-Host "Importing Microsoft.Graph module..." -ForegroundColor Yellow
# Import the Microsoft.Graph module
Import-Module Microsoft.Graph.Authentication
Write-Host "Microsoft.Graph.Authentication module imported successfully."-ForegroundColor Green

# Check if the directory exists
if (-not (Test-Path $ReportingPath -PathType Container)) {
    # If not, create the directory
    New-Item -ItemType Directory -Path $ReportingPath | Out-Null
    Write-Host "Directory created at $ReportingPath"
} else {
    Write-Host "Directory already exists at $ReportingPath"
}

$RequiredScopes = ("DeviceManagementManagedDevices.Read.All", "AuditLog.Read.All", "Directory.Read.All", "User.Read.All", "DeviceManagementApps.Read.All")
Connect-MgGraph -Scope $RequiredScopes -NoWelcome
Write-host "Successfully Connected to MG graph" -ForegroundColor Yellow

Write-Host ""
Write-Host "Step-1: Fetching Device Information from Inutne Portal_Started" -ForegroundColor Yellow

#$Devices = Get-MgDeviceManagementManagedDevice -Filter "OperatingSystem eq 'Windows'" | Select-Object DeviceName, ID, UserDisplayName, UserPrincipalName, UserID, OSVersion,OperatingSystem, Lastsyncdatetime, enrolleddatetime
# Fetch device information based on platform selection

if ($Platform -eq "All") {
    $Devices  = Get-MgDeviceManagementManagedDevice | Select-Object DeviceName, ID, UserDisplayName, UserPrincipalName, UserID, OSVersion,OperatingSystem, Lastsyncdatetime, enrolleddatetime
} else {
    $Devices = Get-MgDeviceManagementManagedDevice -Filter "OperatingSystem eq '$Platform'" | Select-Object DeviceName, ID, UserDisplayName, UserPrincipalName, UserID, OSVersion,OperatingSystem, Lastsyncdatetime, enrolleddatetime
}

$Devices | Export-Csv -Path "$ReportingPath\Devices.csv" -NoTypeInformation
Write-Host "Step-1: Fetching Device Information from Inutne Portal__Completed" -ForegroundColor Green
# Define an array to store the product information
$productArray = @()

# Product information
$productInfo = @"
ProductName, SKUID
$ProductName1 ,$SKUID1
$ProductName2 ,$SKUID2
"@
Write-Host ""
Write-Host "Step-2: Fetching User Account Information from Entra Portal_Started" -ForegroundColor Yellow
# Convert the product information to an array of custom objects
$products = $productInfo | ConvertFrom-Csv
# Extract the SKUIDs from $products
$skuids = $products.SKUID
# Find licensed user accounts
$Headers = @{ ConsistencyLevel = "Eventual" }
$Uri = "https://graph.microsoft.com/beta/users?`$count=true&`$filter=( userType eq 'Member')&$`top=999&`$select=id, mail,accountEnabled,displayName, usertype, signInActivity,assignedLicenses,assignedPlans,assignedLicenses,userPrincipalName,usageLocation,onPremisesSyncEnabled"
[array]$Data = Invoke-MgGraphRequest -Uri $Uri -Headers $Headers
[array]$Users = $Data.Value

If (!($Users)) {
    Write-Host "Not able to find users..."
    break
}
# Paginate until we have all the user accounts
$progress = 0
$totalUsers = $Users.Count
While ($Null -ne $Data.'@odata.nextLink') {
    $Uri = $Data.'@odata.nextLink'
    [array]$Data = Invoke-MgGraphRequest -Uri $Uri -Headers $Headers
    $Users = $Users + $Data.Value
    $progress++
    Write-Progress -Activity "Step-2: Fetching User Account Information From Entra Portal" -Status "Progress" -PercentComplete (($progress / $totalUsers) * 100)
}
Write-Progress -Activity "Fetching User Account Information From Entra Portal" -Completed
Write-Host "Step-2: Fetching User Account Information from Entra Portal_Completed" -ForegroundColor Green
Write-Host ""
Write-Host "Step-3: Processing Per User Specific Information_Started" -ForegroundColor Yellow
# Define progress variables
$totalUsers = $Users.Count
$progressCounter = 0

# Define the initial state of the progress bar
Write-Progress -Activity "Processing Users" -Status "Initializing" -PercentComplete 0

# Create an empty list to store the report
$Report = [System.Collections.Generic.List[Object]]::new()

# Loop through each user
ForEach ($User in $Users) {
    # Update progress counter
    $progressCounter++

    # Calculate progress percentage
    $progressPercentage = ($progressCounter / $totalUsers) * 100

    # Update progress bar
    Write-Progress -Activity "Step-3: Processing Per User Specific Information" -Status "Progress" -PercentComplete $progressPercentage

    # Your existing code for processing each user goes here
    $DaysSinceLastSignIn = $Null
    $DaysSinceLastSuccessfulSignIn = $Null
    $DaysSinceLastSignIn = "N/A"
    $DaysSinceLastSuccessfulSignIn = "N/A"
    $LastSuccessfulSignIn = $User.signInActivity.lastSuccessfulSignInDateTime
    $LastSignIn = $User.signInActivity.lastSignInDateTime
    $accountEnabled = if ($User.accountEnabled -eq $true) { "Account Enabled" } else { "Disabled" }
    $UserMailID = $User.userPrincipalName
    $UserID = $User.ID
    $assignedLicenses = $Users.assignedLicenses.SKUID
    $service = $user.assignedPlans.service
    $usageLocation = $User.usageLocation
    $onPremisesSyncEnabled = if ($user.onPremisesSyncEnabled -eq $true) { "OnPrem Account" } else { "Entra Account" }
    # Extract the SKUIDs from $skuids
    $skuidArray = $skuids -split '\r?\n' | Where-Object { $_ -match '\S' }
    # Extract the servicePlanIds from $user.assignedPlans.servicePlanId
    $servicePlanIds = $user.assignedPlans.servicePlanId -split '\r?\n' | Where-Object { $_ -match '\S' }
    # Check if any common servicePlanIds are assigned to the user
    $servicePlanId1 = "Not Assigned"
    If ($User.assignedLicenses -ne $null) {
        foreach ($servicePlanId in $servicePlanIds) {
            if ($skuidArray -contains $servicePlanId) {
                $servicePlanId1 = "Assigned"
                break
            }
        }
    }
    If (!([string]::IsNullOrWhiteSpace($LastSuccessfulSignIn))) { $DaysSinceLastSuccessfulSignIn = (New-TimeSpan $LastSuccessfulSignIn).Days }
    If (!([string]::IsNullOrWhiteSpace($LastSignIn))) { $DaysSinceLastSignIn = (New-TimeSpan $LastSignIn).Days }

    $LicenceName = $products | Where-Object { $_.SKUID -eq $servicePlanId } | Select-Object -ExpandProperty ProductName

    $DataLine = [PSCustomObject][Ordered]@{
        'User Name' = $User.displayName
        'User Id' = $User.ID
        'User Mani id' = $UserMailID
        'User Location' = $usageLocation
        'Account type' = $onPremisesSyncEnabled
        'Account Enabled' = $accountEnabled
        'Intune Licence' = $servicePlanId1
        'Last successful sign in' = $LastSuccessfulSignIn
        'Last sign in' = $LastSignIn
        'Days since successful sign in' = $DaysSinceLastSuccessfulSignIn
        'Days since sign in' = $DaysSinceLastSignIn
    }
    $Report.Add($DataLine)
}

# Complete progress bar
Write-Progress -Activity "Processing Users" -Status "Complete" -PercentComplete 100
Write-Host "Step-3: Processing Per User Specific Information_Completed" -ForegroundColor Green
Write-Host ""
#$Report | Export-Csv -Path "$ReportingPath\UserAccountDetailedInformation.csv" -NoTypeInformation
# Step-3: Creating Final Report_Started
Write-Host "Step-4: Creating Final Report_Started" -ForegroundColor Yellow

# Define progress variables
$totalDevices = ($Devices.id).Count
$progressCounter = 0
$totalAssignedLicenses = 0
$totalNotAssignedLicenses = 0

# Merging the two reports based on the common field "UserID" with progress bar
$FinalReport = $Devices | ForEach-Object {
    $UserID = $_.UserID
    $UserReport = $Report | Where-Object { $_.'User Id' -eq $UserID }
    $_ | Add-Member -MemberType NoteProperty -Name 'User Location' -Value $UserReport.'User Location'
    $_ | Add-Member -MemberType NoteProperty -Name 'Account Type' -Value $UserReport.'Account type'
    $_ | Add-Member -MemberType NoteProperty -Name 'Account Enabled' -Value $UserReport.'Account Enabled'
    $_ | Add-Member -MemberType NoteProperty -Name 'Intune Licence' -Value $UserReport.'Intune Licence'
    $_ | Add-Member -MemberType NoteProperty -Name 'Last Successful Sign In' -Value $UserReport.'Last successful sign in'
    $_ | Add-Member -MemberType NoteProperty -Name 'Last Sign In' -Value $UserReport.'Last sign in'
    $_ | Add-Member -MemberType NoteProperty -Name 'Days Since Successful Sign In' -Value $UserReport.'Days since successful sign in'
    $_ | Add-Member -MemberType NoteProperty -Name 'Days Since Sign In' -Value $UserReport.'Days since sign in'
    
    # Calculate total assigned and not assigned licenses
    if ($UserReport.'Intune Licence' -eq 'Assigned') {
        $totalAssignedLicenses++
    } elseif ($UserReport.'Intune Licence' -eq 'Not Assigned') {
        $totalNotAssignedLicenses++
    }

    # Increment progress counter
    $progressCounter++
    # Calculate progress percentage
    $progressPercentage = ($progressCounter / $totalDevices) * 100
    # Display progress bar
    Write-Progress -Activity "Step-4 Creating Final Report ($progressCounter / $totalDevices)" -Status "Processing" -PercentComplete $progressPercentage
    # Output the merged report
    $_
}

# Print total device count, total assigned licenses, and total not assigned licenses
Write-Host "Step-4 Creating Final Report_Completed" -ForegroundColor Green
Write-Host ""
Write-Host "Total $Platform Devices In Intune:                         $totalDevices " -ForegroundColor Yellow
Write-Host "Intune Licenses Assigned Count:                            $totalAssignedLicenses" -ForegroundColor Green
Write-Host "Intune Licenses Not Assigned Count:                        $totalNotAssignedLicenses" -ForegroundColor Red
# Display or export the merged report
#$FinalReport | Out-GridView
$FinalReport | Export-Csv -Path "$ReportingPath\Intune_Final_Report.csv" -NoTypeInformation
Write-Host ""
Write-Host "Final report is avaiallbe at $ReportingPath\Intune_Final_Report.csv location " -ForegroundColor Yellow
Disconnect-MgGraph
