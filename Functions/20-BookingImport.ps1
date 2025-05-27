# . ./Functions/14.0-Dictionaries.ps1


$bookingTypes = ($connector.SendRequest([OARequestType]::Read,@{type="BookingType"; method="all"; queryData=@{}; limit=100},$false)).response.Read.BookingType

# Validate if all accounts are on lists
Write-Host "Number of Bookings to create: $($bulkData.Count)"
$dec = Read-Host "Continue (y/N)?"
if($dec.ToLower() -ne "y"){
    Write-Host "Aborted"
    break
}

$validateOnly = $false
. ./Functions/20.1-CheckAllUsers.ps1
. ./Functions/20.2-CheckAllProjects.ps1
# . ./Functions/20.3-ImportBookings.ps1