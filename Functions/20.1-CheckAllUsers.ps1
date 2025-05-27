if(-not $bulkData){
    Write-Host "You need to load data first" -ForegroundColor Red
    break
}

$dataToProcess = $bulkData
$decision = Read-Host "You're about to check $($dataToProcess.Count) Users. Are you sure? (type yes)"
if($decision.ToLower() -ne "yes"){
    Write-Host "User check aborted"
    break
}

$userList = (@($dataToProcess."Resource") + @($dataToProcess."Requester") + @($dataToProcess."Resource Manager")) | Select-Object -Unique

$checkList = @()
$failState = $false
foreach($row in $userList){
    if($name){
        Remove-Variable name
    }
    if($first){
        Remove-Variable first
    }
    if($last){
        Remove-Variable last
    }
    
    if($row -eq "0" -or $row.getType() -eq [OfficeOpenXml.ExcelErrorValue] -or -not $row){
        Write-Warning "User with name $($row) is skipped"
        continue
    }

    $split = $row -split ","
    if ($split.Length -eq 2){
        $first = $split[1].Trim()
        $last = $split[0].Trim()
    }
    else{
        $name = $row.Trim()
    }
    if(-not $first -and -not $last -and -not $name){
        Write-Error "User name is not valid: $($row)"
        $failState = $true
        continue
    }

    $userRead = @{}
    $userRead.type = "User"
    $userRead.method = "equal to"
    $userRead.queryData = @{}
    if($first -and $last){
        $userRead.queryData.first = $first
        $userRead.queryData.last = $last
    }
    if($name){
        $userRead.queryData.name = $name
    }
    $userRead.queryData.active = "1"
    $checkList += $userRead
}

if($failState){
    Write-Error "Cannot progress with update due to errors above"
    break
}

$counter = [pscustomobject] @{ Value = 0 }
$groupSize = 999
$groups = $checkList | Group-Object -Property { [math]::Floor($counter.Value++ / $groupSize) }
if($groupSize -lt $checkList.Count){
    Write-Host "List of users exceeds API limit for one request. Reads are divided to $($groups.Count) groups"
}
foreach($group in $groups){
    $params = @{}
    $params.readData = $group.Group
    if($validateOnly){
        Write-Host "Would send request to check $($group.Group.Count) Users"
        Set-Content -Path ./request.xml -Value ($connector.SendRequest([OARequestType]::ReadBulk,$params,$true)).OuterXml
        continue
    }
    $resp = $connector.SendRequest([OARequestType]::ReadBulk,$params,$false)
    $checkedUsers = $resp.response.Read.User
    Write-Host "Saving transaction file"
    $transactionID = New-Guid
    $successIDs = (($resp.response.Read | Where-Object {$_.status -eq "0"}).User | Select-Object -Property id).id
    $readUsers = ($resp.response.Read.User | Select-Object -Property name).name
    Set-Content -Path "$logFolder/$transactionID.txt" ($successIDs -join ';')
    Set-Content -Path "$logFolder/users-$transactionID.txt" ($readUsers -join ';')
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