# . ./Functions/14.0-Dictionaries.ps1


# Validate if all accounts are on lists
Write-Host "Number of Generics and JobCodes to create: $($bulkData.Count)"
$dec = Read-Host "Continue (y/N)?"
if($dec.ToLower() -ne "y"){
    Write-Host "Aborted"
    break
}

$validateOnly = $false
# . ./Functions/24.1-JobCodeCreate.ps1
# . ./Functions/24.2-GenericUserCreate.ps1
. ./Functions/24.3-AssignGenericsToJC.ps1
. ./Functions/24.4-AssignJCToGenerics.ps1