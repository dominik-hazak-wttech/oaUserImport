if(-not $bulkData){
    Write-Host "You need to load data first" -ForegroundColor Red
    break
}

$dataToProcess = $bulkData
$decision = Read-Host "You're about to create $($dataToProcess.Count) accounts. Are you sure? (type yes)"
if($decision.ToLower() -ne "yes"){
    Write-Host "User account creation aborted"
    break
}
$importList = @()
foreach($row in $dataToProcess){
    if (
        ($row.Generic -split ",").Count -lt 3 -or
        ($row.Generic -split ",").Count -gt 4
       ){
        Write-Error "$($row.Generic) is not a valid generic name. It should be in format: 'Level TechStream, CareerStream, Location'"
        $failState = $true
    }
    $existingGeneric = $row."Generic in OA"
    if($existingGeneric -match "^[0-9]+$"){
        Write-Warning "$($row."Generic") already exists in the system, skipping"
        continue
    }
    $userObj = @{}
    $userObj.parameters = @{}
    $userObj.parameters.nickname = $row.Generic
    $userObj.parameters.name = $row.Generic
    if($row."JobCode in OA" -match "^[0-9]+$"){
        $userObj.parameters.job_codeid = $row."JobCode in OA"
    }
    else{
        if(-not $jobCode){
            Write-Error "ID was not found for '$($row."Job code")' jobcode, leaving field blank."
        }
    }
    $userObj.parameters.active = "1"
    $userObj.parameters.generic = "1"
    $importList += $userObj
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
    $params.usersData = $group.Group
    if($validateOnly){
        Write-Host "Would send request to create group of $($group.Group.Count) accounts"
        Set-Content -Path ./request.xml -Value ($connector.SendRequest([OARequestType]::CreateUserBulk,$params,$true)).OuterXml
        continue
    }
    $resp = $connector.SendRequest([OARequestType]::CreateUserBulk,$params,$false)
    Write-Host "Saving transaction file"
    $transactionID = New-Guid
    $successIDs = (($resp.response.CreateUser | Where-Object {$_.status -eq "0"}).User | Select-Object -Property id).id
    $createdUserIDs = ($resp.response.CreateUser.User | Select-Object -Property nickname).nickname
    Set-Content -Path "$logFolder/$transactionID.txt" ($successIDs -join ';')
    Set-Content -Path "$logFolder/createdUsers-$transactionID.txt" ($createdUserIDs -join ';')
    $failedRequests = [System.Collections.ArrayList]@()
    if ($resp.response.CreateUser.Count -eq 1){
        if($resp.response.CreateUser.status -ne 0){
            $errorResp = $connector.SendRequest([OARequestType]::Read,@{limit="1";type="Error";method="equal to";queryData=@{code=($resp.response.CreateUser[0]).status}})
            $errorInfo = @{
                "First"=$params.usersData.firstName;
                "Last"=$params.usersData.lastName;
                "Error code"=$resp.response.CreateUser.status;
                "Error text"=$errorResp.response.Read.Error.text;
                "OuterXml" = $resp.response.CreateUser.OuterXml
            }
            $failedRequests.Add($errorInfo)
        }
    }
    else{
        for($i=0;$i -lt $resp.response.CreateUser.Count; $i++){
            if(($resp.response.CreateUser[$i]).status -ne 0){
                $errorResp = $connector.SendRequest([OARequestType]::Read,@{limit="1";type="Error";method="equal to";queryData=@{code=($resp.response.CreateUser[$i]).status}})
                $errorInfo = @{
                    First=$params.usersData[$i].firstName;
                    Last=$params.usersData[$i].lastName;
                    "Error code"=($resp.response.CreateUser[$i]).status;
                    "Error text"=$errorResp.response.Read.Error.text;
                    "OuterXml" = $resp.response.CreateUser[$1].OuterXml
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