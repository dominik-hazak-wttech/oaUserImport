# . ./Functions/14.0-Dictionaries.ps1



# Validate if all accounts are on lists
Write-Host "Number of Bookings to create: $($bulkData.Count)"
$dec = Read-Host "Continue (y/N)?"
if($dec.ToLower() -ne "y"){
    Write-Host "Aborted"
    break
}
$bookingTypes = ($connector.SendRequest([OARequestType]::Read,@{type="BookingType"; method="all"; queryData=@{}; limit=100},$false)).response.Read.BookingType

$importToSandbox = $true

Write-Host "Checking users in OpenAir"
. ./Functions/20.1-CheckAllUsers.ps1

Write-Host "Checking projects in OpenAir"
. ./Functions/20.2-CheckAllProjects.ps1

$validateOnly = $true

Write-Host "Importing Bookings to OpenAir"
. ./Functions/20.3-ImportBookings.ps1