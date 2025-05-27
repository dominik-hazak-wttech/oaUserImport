if(-not $bulkData){
    Write-Host "You need to load data first" -ForegroundColor Red
    break
}

$dataToProcess = $bulkData
$userList = (@($dataToProcess."Resource") + @($dataToProcess."Requester") + @($dataToProcess."Resource Manager")) | Select-Object -Unique

# $decision = Read-Host "You're about to check $($userList.Count) Users. Are you sure? (type yes)"
# if($decision.ToLower() -ne "yes"){
#     Write-Host "User check aborted"
#     break
# }

$skipped=0
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
        $skipped++
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

if($skipped -gt 0){
    Write-Host "$skipped entries were skipped"
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
    Write-Host "Transaction ID: $transactionID"
    $successIDs = (($resp.response.Read | Where-Object {$_.status -eq "0"}).User | Select-Object -Property id).id
    $readUsers = ($resp.response.Read.User | Select-Object -Property name).name
    Set-Content -Path "$logFolder/$transactionID.txt" ($successIDs -join ';')
    Set-Content -Path "$logFolder/users-$transactionID.txt" ($readUsers -join ';')
    Write-Host "Reading errors if any"
    $failedRequests = [System.Collections.ArrayList]@()
    if ($resp.response.Read.Count -eq 1){
        if($resp.response.Read.status -ne 0){
            $errorResp = $connector.SendRequest([OARequestType]::Read,@{limit="1";type="Error";method="equal to";queryData=@{code=($resp.response.Read[0]).status}})
            $errorInfo = @{}
            $errorInfo["Error code"]=$resp.response.Add.status;
            $errorInfo["Error text"]=$errorResp.response.Read.Error.comment;
            $errorInfo["OuterXml"] = $resp.response.Read.OuterXml
            if($params.readData.queryData.name){
                $errorInfo["Name"]=$params.readData.queryData.name;
            }
            if($params.readData.queryData.first -and $params.readData.queryData.last){
                $errorInfo["Name"]="$($params.readData.queryData.first) $($params.readData.queryData.last)";
            }

            $failedRequests.Add($errorInfo) | Out-Null
        }
    }
    else{
        for($i=0;$i -lt $resp.response.Read.Count; $i++){
            if(($resp.response.Read[$i]).status -ne 0){
                $errorResp = $connector.SendRequest([OARequestType]::Read,@{limit="1";type="Error";method="equal to";queryData=@{code=($resp.response.Read[$i]).status}})
                $errorInfo = @{}
                $errorInfo["Error code"]=$resp.response.Read[$i].status;
                $errorInfo["Error text"]=$errorResp.response.Read.Error.comment;
                $errorInfo["OuterXml"]= $resp.response.Read[$i].OuterXml
                if($params.readData[$i].queryData.name){
                    $errorInfo["Name"]=$params.readData[$i].queryData.name;
                }
                if($params.readData[$i].queryData.first -and $params.readData[$i].queryData.last){
                    $errorInfo["Name"]="$($params.readData[$i].queryData.first) $($params.readData[$i].queryData.last)";
                }
                $failedRequests.Add($errorInfo) | Out-Null
            }
        }
    }
    Set-Content -Path "$logFolder/error-$transactionID.txt" ($failedRequests | Format-Table | Out-String)
    Write-Host "Out of $($params.readData.Count):`n`t$($successIDs.Count) were read successfully`n`t$($failedRequests.Count) failed"
    Set-Content -Path "$logFolder/response-$transactionID.xml" $resp.response.OuterXml
}