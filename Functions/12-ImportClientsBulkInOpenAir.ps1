if(-not $bulkData){
    Write-Host "You need to load data first" -ForegroundColor Red
    break
}
$dataToProcess = $bulkData
$decision = Read-Host "You're about to add $($dataToProcess.Count) clients. Are you sure? (type yes)"
if($decision.ToLower() -ne "yes"){
Write-Host "Client creation aborted"
    break
}
$importList = @()
foreach($row in $dataToProcess){
    Write-Host $row."Client name"
    $clientObj = @{}
    $clientObj.type = "Customer"
    $clientObj.dataToAdd = @{}
    $clientObj.dataToAdd.name = $row."Client name"
    $clientObj.dataToAdd.company = $row."Client name"
    $clientObj.dataToAdd.userid = $row."Client Owner ID"
    $clientObj.dataToAdd.Portfolio__c = $row."Portfolio"
    $clientObj.dataToAdd.Client_code__c = $row."OA code to be created"
    $importList += $clientObj
}
$counter = [pscustomobject] @{ Value = 0 }
$groupSize = 999
$groups = $importList | Group-Object -Property { [math]::Floor($counter.Value++ / $groupSize) }
if($groupSize -lt $importList.Count){
    Write-Host "List of users exceeds API limit for one request. Users are divided to $($groups.Count) groups"
}
foreach($group in $groups){
    $params = @{}
    $params.addRequests = $group.Group
    Write-Host $params.addRequests | Out-String
    $resp = $connector.SendRequest([OARequestType]::AddBulk,$params)
    Write-Host "Saving transaction file"
    $transactionID = New-Guid
    $successIDs = (($resp.response.Add | Where-Object {$_.status -eq "0"}).Customer | Select-Object -Property id).id
    $createdClientIDs = ($resp.response.Add.Customer | Select-Object -Property name).name
    Set-Content -Path "$logFolder/$transactionID.txt" ($successIDs -join ';')
    Set-Content -Path "$logFolder/addedClients-$transactionID.txt" ($createdClientIDs -join ';')
    $failedRequests = @()
    if ($resp.response.Add.Count -eq 1){
        if($resp.response.Add.status -ne 0){
            $errorResp = $connector.SendRequest([OARequestType]::Read,@{limit="1";type="Error";method="equal to";queryData=@{code=($resp.response.Add[$i]).status}})
            $failedRequests += @{
                name=$params.addRequests[$i].name;
                "Error code"=($resp.response.Add[$i]).status;
                "Error text"=$errorResp.response.Read.Error.text;
                "OuterXml" = $resp.response.Add[$1].OuterXml
            }
        }
    }
    else{
        for($i=0;$i -lt $resp.response.Add.Count; $i++){
            if(($resp.response.Add[$i]).status -ne 0){
                $errorResp = $connector.SendRequest([OARequestType]::Read,@{limit="1";type="Error";method="equal to";queryData=@{code=($resp.response.Add[$i]).status}})
                $failedRequests += @{
                    name=$params.addRequests[$i].name;
                    "Error code"=($resp.response.Add[$i]).status;
                    "Error text"=$errorResp.response.Read.Error.text;
                    "OuterXml" = $resp.response.Add[$1].OuterXml
                }
            }
        }
    }
    Set-Content -Path "$logFolder/error-$transactionID.txt" ($failedRequests | Format-List | Out-String)
    Write-Host "Out of $($params.addRequests.Count):`n`t$($successIDs.Count) were created successfully`n`t$($failedRequests.Count) failed"
    Write-Host "Transaction ID: $transactionID"
    Set-Content -Path "$logFolder/response-$transactionID.xml" $resp.response.OuterXml
}