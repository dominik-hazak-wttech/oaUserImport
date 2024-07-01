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

$dataToProcess = $bulkData
$dataToProcess = $dataToProcess | Where-Object {$_."User in OA?" -ne ""}
$dataToProcess = $dataToProcess | Where-Object {$null -ne $_."User in OA?"}
$dataToProcess = $dataToProcess | Where-Object {$_."OA Import Status" -ne $OAStatusSkip}

Write-Host "Reading project groups"
$groupResp = $connector.SendRequest([OARequestType]::Read,@{"type"="Projectgroup";"method"="all";"limit"=100;"queryData"=@{"name" = ""}},$false)
$projectgroups = $groupResp.response.Read.Projectgroup

Write-Host "Done reading"
Write-Host "Found $(($projectgroups | Measure-Object).Count) groups:"
foreach($projectgroup in $projectgroups){
    Write-Host "`t* $($projectgroup.name)"
}

$decision = Read-Host "You're about to udpate $($dataToProcess.Count) accounts. Are you sure? (type yes)"
if($decision.ToLower() -ne "yes"){
    Write-Host "User account creation aborted"
    break
}
$groupModifys = @{}
foreach($row in $dataToProcess){
    if ($row."User ID" -eq "" -or $null -eq $row."User ID"){
        Write-Error "$($row.Email) is missing User ID value!"
        $failState = $true
    }
    if (-not $userDataDictionary[$row."User ID"]){
        Write-Warning "$($row."User ID") is missing id in dictionary! User will be skipped!"
        continue
    }
    
    $userGroups = $row."Assignment groups" -split " \| "
    $userId = $userDataDictionary[$row."User ID"]
    foreach($groupName in $userGroups){
        if($groupModifys.ContainsKey($groupName)){
            if($groupModifys[$groupName].dataToUpdate.assigned_users -notcontains "*$userId*"){
                $groupModifys[$groupName].dataToUpdate.assigned_users = $groupModifys[$groupName].dataToUpdate.assigned_users + $userId
            }
        }
        else{
            $groupObject = $projectgroups | Where-Object {$_.name -eq $groupName}
            if(-not $groupObject){
                Write-Warning "Group '$groupName' does not exist! This group will be skipped!"
                continue
            }
            $groupModifys[$groupName] = @{
                "id" = $groupObject.id;
                "type" = "Projectgroup";
                "dataToUpdate" = @{
                    "assigned_users" = ($groupObject.assigned_users -split ",") + $userId
                }
            }
        }
    }
}
$global:groupModifys = $groupModifys
foreach($groupName in $groupModifys.Keys){
    $groupModifys[$groupName].dataToUpdate.assigned_users = $groupModifys[$groupName].dataToUpdate.assigned_users -join ","
}
# $groupModifys.dataToUpdate.assigned_users
# break
$validateOnly=$true


if($failState){
    Write-Error "Cannot progress with update due to errors above"
    break
}

$importList = $groupModifys.Values

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
        Write-Host "Would send request to edit $($group.Group.Count) groups"
        Set-Content -Path ./request.xml -Value ($connector.SendRequest([OARequestType]::ModifyBulk,$params,$true)).OuterXml
        continue
    }
    $resp = $connector.SendRequest([OARequestType]::ModifyBulk,$params,$false)
    Write-Host "Saving transaction file"
    $transactionID = New-Guid
    $successIDs = (($resp.response.Modify | Where-Object {$_.status -eq "0"}).User | Select-Object -Property id).id
    $modifiedUserNames = ($resp.response.Modify.User | Select-Object -Property nickname).nickname
    Set-Content -Path "$logFolder/$transactionID.txt" ($successIDs -join ';')
    Set-Content -Path "$logFolder/modifiedUsers-$transactionID.txt" ($modifiedUserNames -join ';')
    $failedRequests = [System.Collections.ArrayList]@()
    if ($resp.response.Modify.Count -eq 1){
        if($resp.response.Modify.status -ne 0){
            $errorResp = $connector.SendRequest([OARequestType]::Read,@{limit="1";type="Error";method="equal to";queryData=@{code=($resp.response.Modify[$i]).status}})
            $errorInfo = @{
                "ID" = $params.modifyRequests.id;
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
                    "ID" = $params.modifyRequests[$i].id;
                    "Error code"=($resp.response.Modify[$i]).status;
                    "Error text"=$errorResp.response.Read.Error.text;
                    "OuterXml" = ($resp.response.Modify[$1]).OuterXml
                }
                $failedRequests.Add($errorInfo)
            }
        }
    }
    Set-Content -Path "$logFolder/error-$transactionID.txt" ($failedRequests | Format-List | Out-String)
    Write-Host "Out of $($params.usersData.Count):`n`t$($successIDs.Count) were modified successfully`n`t$($failedRequests.Count) failed"
    Write-Host "Transaction ID: $transactionID"
    Set-Content -Path "$logFolder/response-$transactionID.xml" $resp.OuterXml
}