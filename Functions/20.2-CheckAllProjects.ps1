if(-not $bulkData){
    Write-Host "You need to load data first" -ForegroundColor Red
    break
}

$dataToProcess = $bulkData
$projectList = $dataToProcess."Project" | Select-Object -Unique

# $decision = Read-Host "You're about to check $($projectList.Count) Projects. Are you sure? (type yes)"
# if($decision.ToLower() -ne "yes"){
#     Write-Host "Project check aborted"
#     break
# }

$skipped=0
$checkList = @()
$failState = $false
foreach($row in $projectList){
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
        Write-Warning "Project with name $($row) is skipped"
        $skipped++
        continue
    }
    
    $name = $row.Trim()
    if(-not $name){
        Write-Error "Project name is not valid: $($row)"
        $failState = $true
        continue
    }

    $projectRead = @{}
    $projectRead.type = "Project"
    $projectRead.method = "equal to"
    $projectRead.queryData = @{}
    $projectRead.queryData.name = $name
    $checkList += $projectRead
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
        Write-Host "Would send request to check $($group.Group.Count) Projects"
        Set-Content -Path ./request.xml -Value ($connector.SendRequest([OARequestType]::ReadBulk,$params,$true)).OuterXml
        continue
    }
    $resp = $connector.SendRequest([OARequestType]::ReadBulk,$params,$false)
    $checkedProjects = $resp.response.Read.Project
    Write-Host "Saving transaction file"
    $transactionID = New-Guid
    Write-Host "Transaction ID: $transactionID"
    $successIDs = (($resp.response.Read | Where-Object {$_.status -eq "0"}).Project | Select-Object -Property id).id
    $readProjects = $resp.response.Read.Project | Select-Object -Property id,name
    Set-Content -Path "$logFolder/$transactionID.json" ($successIDs | ConvertTo-Json)
    Set-Content -Path "$logFolder/projects-$transactionID.json" ($readProjects | ConvertTo-Json)
    $failedRequests = [System.Collections.ArrayList]@()
    if ($resp.response.Read.Count -eq 1){
        if($resp.response.Read.status -ne 0){
            $errorResp = $connector.SendRequest([OARequestType]::Read,@{limit="1";type="Error";method="equal to";queryData=@{code=($resp.response.Read[0]).status}})
            $errorInfo = @{
                "Name"=$params.readData.queryData.name;
                "Error code"=$resp.response.Read.status;
                "Error text"=$errorResp.response.Read.Error.comment;
                "OuterXml" = $resp.response.Read.OuterXml
            }
            $failedRequests.Add($errorInfo) | Out-Null
        }
    }
    else{
        for($i=0;$i -lt $resp.response.Read.Count; $i++){
            if(($resp.response.Read[$i]).status -ne 0){
                $errorResp = $connector.SendRequest([OARequestType]::Read,@{limit="1";type="Error";method="equal to";queryData=@{code=($resp.response.Read[$i]).status}})
                $errorInfo = @{
                    "Name"=$params.readData[$i].queryData.name;
                    "Error code"=($resp.response.Read[$i]).status;
                    "Error text"=$errorResp.response.Read.Error.comment;
                    "OuterXml" = $resp.response.Read[$i].OuterXml
                }
                $failedRequests.Add($errorInfo) | Out-Null
            }
        }
    }
    Set-Content -Path "$logFolder/error-$transactionID.json" ($failedRequests | ConvertTo-Json | Out-String)
    Write-Host "Out of $($params.readData.Count):`n`t$($successIDs.Count) were read successfully`n`t$($failedRequests.Count) failed"
    Set-Content -Path "$logFolder/response-$transactionID.xml" $resp.response.OuterXml
}