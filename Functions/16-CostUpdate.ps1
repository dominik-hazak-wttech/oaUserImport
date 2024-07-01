if(-not $bulkData){
    Write-Host "You need to load data first" -ForegroundColor Red
    break
}

$OAStatusReady = "READY FOR IMPORT"
$OACleaningStatusValidate = "NEEDS VALIDATION"
$OAStatusImported = "IMPORTED"
$OAStatusSkip = "SKIP"
$OARegularAccount = "Regular"
$OASkeletonAccount = "Skeleton"

. ./Functions/14.0-Dictionaries.ps1

function ConvertTo-Datetime($oaDate){
    return (Get-Date -Year $oaDate.Date.year -Month $oaDate.Date.month -Day $oaDate.Date.day)
}

$dataToProcess = $bulkData
$dataToProcess = $dataToProcess | Where-Object {$_."User in OA?" -ne ""}
$dataToProcess = $dataToProcess | Where-Object {$null -ne $_."User in OA?"}
$dataToProcess = $dataToProcess | Where-Object {$_."OA Import Status" -ne $OAStatusSkip}
$dataToProcess = $dataToProcess | Where-Object {$_."IMPORT_Cost" -eq "yes"}

$decision = Read-Host "You're about to udpate $($dataToProcess.Count) costs. Are you sure? (type yes)"
if($decision.ToLower() -ne "yes"){
    Write-Host "User account's cost update aborted"
    break
}

$transactionID = New-Guid
# $validateOnly = $true
$totalDataCount = $dataToProcess.Count
$i=0
Write-Host "Current transaction id: $transactionID"
Write-Progress -Activity "Updating costs" -Id 1 -PercentComplete 0
foreach($row in $dataToProcess){
    if ($row."User ID" -eq "" -or $null -eq $row."User ID"){
        Write-Error "$($row.Email) is missing User ID value!"
        $failState = $true
    }
    if (-not $userDataDictionary[$row."User ID"]){
        Write-Warning "$($row."User ID") is missing id in dictionary! User will be skipped!"
        continue
    }
    
    $costUpdate = @{}
    $userId = $userDataDictionary[$row."User ID"]
    $userCost = $row."Cost"
    $currentCostsInOA = $connector.SendRequest([OARequestType]::Read,@{"type"="LoadedCost";"method"="equal to";"limit"=100;"queryData"=@{"userid" = "$userId"}},$false)
    $activeCost = ($currentCostsInOA.response.Read | Where-Object {$_.status -eq 0}).LoadedCost | Where-Object {$_.current -eq "1"}
    $otherCosts = ($currentCostsInOA.response.Read | Where-Object {$_.status -eq 0}).LoadedCost | Where-Object {$_.current -eq "0"}
    $newestCost = Get-Date -Day 1 -Month 1 -Year 2019 
    foreach($cost in $otherCosts){
        $dateConverted = ConvertTo-Datetime($cost.end)
        if($dateConverted -gt $newestCost){
            $newestCost = $dateConverted
        }
    }
    $newestCost = $newestCost.AddDays(1)
    if($activeCost){
        $costUpdate.Update = @{}
        $costUpdate.Update.id = $activeCost.id
        $costUpdate.Update.dataToUpdate = @{}
        $costUpdate.Update.dataToUpdate.current = "0"
        $costUpdate.Update.dataToUpdate.end = @{"year"="2024";"month"="06";"day"="30"}
        $costUpdate.Update.dataToUpdate.start = @{"year"=$newestCost.year;"month"=$newestCost.month;"day"=$newestCost.day}
    }
    else {
        Write-Warning "No active cost for user $($row."User ID")"
    }
    $costUpdate.Add = @{}
    $costUpdate.Add.dataToAdd = @{}
    $costUpdate.Add.dataToAdd.cost = $userCost
    $costUpdate.Add.dataToAdd.userid = $userId
    $costUpdate.Add.dataToAdd.currency = "GBP"
    $costUpdate.Add.dataToAdd.current = "1"

    if($validateOnly){
        Write-Host "Would send request to cost for $($row."User ID")"
        Add-Content -Path ./request.xml -Value ($connector.SendRequest([OARequestType]::CostUpdate,@{costObjects=$costUpdate},$true)).OuterXml
        continue
    }
    $resp = $connector.SendRequest([OARequestType]::CostUpdate,@{costObjects=$costUpdate},$false)
    Add-Content -Path "$logFolder/response-$transactionID.xml" $resp.OuterXml
    Add-Content -Path "$logFolder/updatedCosts-$transactionID.xml" $resp.response.Update.LoadedCost.id
    Add-Content -Path "$logFolder/addedCosts-$transactionID.xml" $resp.response.Add.LoadedCost.id
    Write-Progress -id 1 -Activity "Updating costs" -PercentComplete ($i*100/$totalDataCount)
    $i++
}