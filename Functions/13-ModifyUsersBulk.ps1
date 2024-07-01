if(-not $bulkData){
    Write-Host "You need to load data first" -ForegroundColor Red
    break
}

$dataToProcess = $bulkData
# $dataToProcess = $dataToProcess | Where-Object {$_."OA Import Status" -eq "READY FOR IMPORT" -and $_."User in OA?" -eq "MATCH"}
$decision = Read-Host "You're about to modify $($dataToProcess.Count) accounts. Are you sure? (type yes)"
if($decision.ToLower() -ne "yes"){
    Write-Host "User account creation aborted"
    break
}

Write-Host "Reading user info"
$importList = @()
foreach($row in $dataToProcess){
    $userObj = @{}
    $userObj.id = $row.id
    $userObj.type = "User"
    $userObj.dataToUpdate = @{}
    $userObj.dataToUpdate.Company__c = $row.custom_55
    $importList += $userObj
}

$counter = [pscustomobject] @{ Value = 0 }
$groupSize = 999
$groups = $importList | Group-Object -Property { [math]::Floor($counter.Value++ / $groupSize) }
if($groupSize -lt $importList.Count){
    Write-Host "List of users exceeds API limit for one request. Users are divided to $($groups.Count) groups"
}

foreach($group in $groups){
    $params = @{}
    $params.modifyRequests = $group.Group
    $resp = $connector.SendRequest([OARequestType]::ModifyBulk,$params,$false)
    Write-Host "Saving transaction file"
    $transactionID = New-Guid
    $successIDs = (($resp.response.Modify | Where-Object {$_.status -eq "0"}).User | Select-Object -Property id).id
    $modifiedUserNames = ($resp.response.Modify.User | Select-Object -Property nickname).nickname
    Set-Content -Path "$logFolder/$transactionID.txt" ($successIDs -join ';')
    Set-Content -Path "$logFolder/modifiedUsers-$transactionID.txt" ($modifiedUserNames -join ';')
    $failedRequests = @()
    if ($resp.response.Modify.Count -eq 1){
        if($resp.response.Modify.status -ne 0){
            $errorResp = $connector.SendRequest([OARequestType]::Read,@{limit="1";type="Error";method="equal to";queryData=@{code=($resp.response.Modify[$i]).status}})
            $failedRequests += @{
                ID = $params.modifyRequests.id;
                "Error code"=($resp.response.Modify[$i]).status;
                "Error text"=$errorResp.response.Read.Error.text;
                "OuterXml" = $resp.response.Modify[$1].OuterXml
            }
        }
    }
    else{
        for($i=0;$i -lt $resp.response.Modify.Count; $i++){
            if(($resp.response.Modify[$i]).status -ne 0){
                $errorResp = $connector.SendRequest([OARequestType]::Read,@{limit="1";type="Error";method="equal to";queryData=@{code=($resp.response.Modify[$i]).status}})
                $failedRequests += @{
                    ID = $params.modifyRequests[$i].firstName;
                    "Error code"=($resp.response.Modify[$i]).status;
                    "Error text"=$errorResp.response.Read.Error.text;
                    "OuterXml" = $resp.response.Modify[$1].OuterXml
                }
            }
        }
    }
    Set-Content -Path "$logFolder/error-$transactionID.txt" ($failedRequests | Format-List | Out-String)
    Write-Host "Out of $($params.usersData.Count):`n`t$($successIDs.Count) were modified successfully`n`t$($failedRequests.Count) failed"
    Write-Host "Transaction ID: $transactionID"
    Set-Content -Path "$logFolder/response-$transactionID.xml" $resp.OuterXml
}