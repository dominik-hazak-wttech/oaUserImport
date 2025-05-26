if(-not $bulkData){
    Write-Host "You need to load data first" -ForegroundColor Red
    break
}

$dataToProcess = $bulkData
$decision = Read-Host "You're about to assign $($dataToProcess.Count) accounts to job codes. Are you sure? (type yes)"
if($decision.ToLower() -ne "yes"){
    Write-Host "Operation aborted"
    break
}
$importList = @()
foreach($row in $dataToProcess){
    $generic = $row."Assigned generic"
    $jobCode = $row."JobCode in OA"
    if (-not ($generic -match "^[0-9]+$")){
        Write-Error "ID was not found for '$($row.Generic)' generic."
        $failState = $true
    }
    if(-not ($jobCode -match "^[0-9]+$")){
        Write-Error "ID was not found for '$($row."Job Code")' jobcode."
        $failState = $true
    }
    $jcObj = @{}
    $jcObj.id = $jobCode
    $jcObj.type = "Jobcode"
    $jcObj.dataToUpdate = @{}
    $jcObj.dataToUpdate.userid_fte = $generic
    $importList += $jcObj
}

if($failState){
    Write-Error "Cannot progress with update due to errors above"
    break
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
    if($validateOnly){
        Write-Host "Would send request to create group of $($group.Group.Count) accounts"
        Set-Content -Path ./request.xml -Value ($connector.SendRequest([OARequestType]::ModifyBulk,$params,$true)).OuterXml
        continue
    }
    $resp = $connector.SendRequest([OARequestType]::ModifyBulk,$params,$false)
    Write-Host "Saving transaction file"
    $transactionID = New-Guid
    $successIDs = (($resp.response.Modify | Where-Object {$_.status -eq "0"}).User | Select-Object -Property id).id
    $createdUserIDs = ($resp.response.Modify.User | Select-Object -Property nickname).nickname
    Set-Content -Path "$logFolder/$transactionID.txt" ($successIDs -join ';')
    Set-Content -Path "$logFolder/modified-$transactionID.txt" ($createdUserIDs -join ';')
    $failedRequests = [System.Collections.ArrayList]@()
    if ($resp.response.Modify.Count -eq 1){
        if($resp.response.Modify.status -ne 0){
            $errorResp = $connector.SendRequest([OARequestType]::Read,@{limit="1";type="Error";method="equal to";queryData=@{code=($resp.response.Modify[$i]).status}})
            $errorInfo = @{
                "First"=$params.usersData.firstName;
                "Last"=$params.usersData.lastName;
                "Error code"=$resp.response.Modify.status;
                "Error text"=$errorResp.response.Read.Error.text;
                "OuterXml" = $resp.response.Modify.OuterXml
            }
            $failedRequests.Add($errorInfo)
        }
    }
    else{
        for($i=0;$i -lt $resp.response.Modify.Count; $i++){
            if(($resp.response.Modify[$i]).status -ne 0){
                $errorResp = $connector.SendRequest([OARequestType]::Read,@{limit="1";type="Error";method="equal to";queryData=@{code=($resp.response.Modify[$i]).status}})
                $errorInfo = @{
                    First=$params.usersData[$i].firstName;
                    Last=$params.usersData[$i].lastName;
                    "Error code"=($resp.response.Modify[$i]).status;
                    "Error text"=$errorResp.response.Read.Error.text;
                    "OuterXml" = $resp.response.Modify[$1].OuterXml
                }
                $failedRequests.Add($errorInfo)
            }
        }
    }
    Set-Content -Path "$logFolder/error-$transactionID.txt" ($failedRequests | Format-List | Out-String)
    Write-Host "Out of $($params.usersData.Count):`n`t$($successIDs.Count) were created successfully`n`t$($failedRequests.Count) failed"
    Write-Host "Transaction ID: $transactionID"
    Set-Content -Path "$logFolder/response-$transactionID.xml" $resp.response.OuterXml
}