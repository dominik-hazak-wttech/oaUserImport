# if(-not $bulkData){
#     Write-Host "You need to load data first" -ForegroundColor Red
#     break
# }

. ./Functions/14.0-Dictionaries.ps1

$OAStatusReady = "READY FOR IMPORT"
$OACleaningStatusValidate = "NEEDS VALIDATION"
$TD_NoMatch = "NO MATCH"
$TD_Match = "MATCH"

$usersToCreate = $bulkData | Where-Object {$_."User in OA?" -eq $TD_NoMatch}
$usersToUpdate = $bulkData | Where-Object {$_."User in OA?" -eq $TD_Match}
$usersNotSelected = $bulkData | Where-Object {$_."User in OA?" -ne $TD_NoMatch -and $_."User in OA?" -ne $TD_Match}

# Validate if all accounts are on lists
if ($bulkData.Count -ne ($usersToCreate.Count + $usersToUpdate.Count)){
    Write-Error "Not all accounts are split in one list or another! `nBulk data has $($bulkData.Count) accounts but lists have $($usersToCreate.Count + $usersToUpdate.Count) accounts!"
    Write-Host "User IDs:"
    Write-Host ($usersNotSelected."User ID" | Format-List | Out-String)
    $cont = Read-Host "Continue anyway? (y/n)"
    if($cont.ToLower() -ne "y"){
        break
    }
}

Write-Host "Users to create: $($usersToCreate.Count)"
Write-Host "Users to update: $($usersToUpdate.Count)"
$dec = Read-Host "Continue (y/N)?"
if($dec.ToLower() -ne "y"){
    Write-Host "Aborted"
    break
}
# $validateOnly = $true
# . ./Functions/14.2-UserCreate.ps1
# . ./Functions/14.1-UserUpdate.ps1
. ./Functions/14.3-UserUpdate-fixLogin.ps1