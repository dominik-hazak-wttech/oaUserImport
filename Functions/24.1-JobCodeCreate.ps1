if(-not $bulkData){
    Write-Host "You need to load data first" -ForegroundColor Red
    break
}

$dataToProcess = $bulkData
$decision = Read-Host "You're about to create $($dataToProcess.Count) JobCodes. Are you sure? (type yes)"
if($decision.ToLower() -ne "yes"){
    Write-Host "Job code creation aborted"
    break
}
$importList = @()
foreach($row in $dataToProcess){
    $split = $row."Job Code" -split "_"
    if ($split.Count -ne 2){
        Write-Error "$($row."Job Code") is not a valid job code. It should be in format: 'VaultCode_Loc'"
        $failState = $true
    }
    if ($split[0] -ne $row."Vault Code Proposal"){
        Write-Error "$($row."Job Code"): first part does not match Vault code proposal"
        $failState = $true
    }
    if (-not ($row."Location2" -like "$($split[1])*")){
        Write-Error "$($row."Job Code"): second part does not match with location"
        $failState = $true
    }
    $existingJobCode = $row."JobCode in OA"
    if($existingJobCode -match "^[0-9]+$"){
        Write-Warning "$($row."Job Code") already exists with id $($row."JobCode in OA") in the system, skipping"
        continue
    }
    $jcObj = @{}
    $jcObj.type = "Jobcode"
    $jcObj.dataToAdd = @{}
    $jcObj.dataToAdd.name = $row."Job Code"
    $jcObj.dataToAdd.active = "1"
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
    Write-Host "List of users exceeds API limit for one request. JobCodes are divided to $($groups.Count) groups"
}
foreach($group in $groups){
    $params = @{}
    $params.addRequests = $group.Group
    if($validateOnly){
        Write-Host "Would send request to create $($group.Group.Count) JobCodes"
        Set-Content -Path ./request.xml -Value ($connector.SendRequest([OARequestType]::AddBulk,$params,$true)).OuterXml
        continue
    }
    $resp = $connector.SendRequest([OARequestType]::AddBulk,$params,$false)
    Write-Host "Saving transaction file"
    $transactionID = New-Guid
    $successIDs = (($resp.response.Add | Where-Object {$_.status -eq "0"}).Jobcode | Select-Object -Property id).id
    $createdJCs = ($resp.response.Add.Jobcode | Select-Object -Property name).name
    Set-Content -Path "$logFolder/$transactionID.txt" ($successIDs -join ';')
    Set-Content -Path "$logFolder/createdJobcodes-$transactionID.txt" ($createdJCs -join ';')
    $failedRequests = [System.Collections.ArrayList]@()
    if ($resp.response.Add.Count -eq 1){
        if($resp.response.Add.status -ne 0){
            $errorResp = $connector.SendRequest([OARequestType]::Read,@{limit="1";type="Error";method="equal to";queryData=@{code=($resp.response.Add[0]).status}})
            $errorInfo = @{
                "Name"=$params.addRequests.name;
                "Error code"=$resp.response.Add.status;
                "Error text"=$errorResp.response.Read.Error.text;
                "OuterXml" = $resp.response.Add.OuterXml
            }
            $failedRequests.Add($errorInfo)
        }
    }
    else{
        for($i=0;$i -lt $resp.response.Add.Count; $i++){
            if(($resp.response.Add[$i]).status -ne 0){
                $errorResp = $connector.SendRequest([OARequestType]::Read,@{limit="1";type="Error";method="equal to";queryData=@{code=($resp.response.Add[$i]).status}})
                $errorInfo = @{
                    Name=$params.addRequests[$i].name;
                    "Error code"=($resp.response.Add[$i]).status;
                    "Error text"=$errorResp.response.Read.Error.text;
                    "OuterXml" = $resp.response.Add[$1].OuterXml
                }
                $failedRequests.Add($errorInfo)
            }
        }
    }
    Set-Content -Path "$logFolder/error-$transactionID.txt" ($failedRequests | Format-List | Out-String)
    Write-Host "Out of $($params.addRequests.Count):`n`t$($successIDs.Count) were created successfully`n`t$($failedRequests.Count) failed"
    Write-Host "Transaction ID: $transactionID"
    Set-Content -Path "$logFolder/response-$transactionID.xml" $resp.response.OuterXml
}