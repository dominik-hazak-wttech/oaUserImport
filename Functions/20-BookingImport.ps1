# . ./Functions/14.0-Dictionaries.ps1

$rateLevels = @{
        "UK" = "Onshore";
        "NL" = "Onshore";
        "DK" = "Onshore";
        "BE" = "Onshore";
        "US" = "Onshore";
        "AU" = "Onshore";
        "PL" = "Nearshore";
        "PT" = "Nearshore";
        "HU" = "Nearshore";
        "LT" = "Nearshore";
        "GR" = "Nearshore";
        "IN" = "Offshore";
        "RU" = "Offshore";
        "SA" = "Offshore";
        "CN" = "Offshore";
        "OTH" = "Offshore"
}

$capabilities = @{
    "Business Consultancy" = "Business Services";
    "Business Services Management" = "Business Services";
    "Change & Enablement" = "Business Services";
    "Content" = "Business Services";
    "CX Consultancy" = "Business Services";
    "CXM" = "Business Services";
    "Feed Management & Intelligence" = "Business Services";
    "Implementation" = "Business Services";
    "Insights" = "Business Services";
    "Optimisation" = "Business Services";
    "Performance Marketing" = "Business Services";
    "Product Design - UI" = "Business Services";
    "Product Design - UX" = "Business Services";
    "Product Management" = "Business Services";
    "Retail Media" = "Business Services";
    "Social Commerce" = "Business Services";
    "Technical Consultancy" = "Business Services";
    "UI" = "Business Services";
    "UX" = "Business Services";
    "Client Services" = "Client Services";
    "Business Analysis" = "Delivery";
    "Help Desk" = "Delivery";
    "Project Management" = "Delivery";
    "Service Management" = "Delivery";
    "Service Operations" = "Delivery";
    "Service Technical" = "Delivery";
    "Technical Service Management" = "Delivery";
    "AI Engineering" = "Engineering";
    "App (Engineering)" = "Engineering";
    "Application Engineering" = "Engineering";
    "Back-end Engineering" = "Engineering";
    "Content (Engineering)" = "Engineering";
    "Engineering Management" = "Engineering";
    "Front-end Architecture" = "Engineering";
    "Front-end Engineering" = "Engineering";
    "Performance Engineering" = "Engineering";
    "Platform Engineering" = "Engineering";
    "Project QA Management" = "Engineering";
    "QA Consultancy" = "Engineering";
    "QA Engineering" = "Engineering";
    "Technical Leadership" = "Engineering";
    "Other" = "Other";
    "Solution Architecture" = "Technology";
    "Technical Architecture" = "Technology";
    "Technology Management" = "Technology"
}

# Validate if all accounts are on lists
Write-Host "Number of Bookings to process: $($bulkData.Count)"
$dec = Read-Host "Continue (y/N)?"
if($dec.ToLower() -ne "y"){
    Write-Host "Aborted"
    break
}
$bookingTypes = ($connector.SendRequest([OARequestType]::Read,@{type="BookingType"; method="all"; queryData=@{}; limit=100},$false)).response.Read.BookingType

# $importToSandbox = $true
$validateOnly = $false

Write-Host "Checking users in OpenAir"
. ./Functions/20.1-CheckAllUsers.ps1

Write-Host "Checking projects in OpenAir"
. ./Functions/20.2-CheckAllProjects.ps1

# $validateOnly = $true

Write-Host "Importing Bookings to OpenAir"
. ./Functions/20.3-ImportBookings.ps1

Write-Host "Reading Bookings from OpenAir"
. ./Functions/20.4-ReadBookings.ps1