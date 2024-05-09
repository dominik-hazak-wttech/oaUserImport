if(-not $bulkData){
    Write-Host "You need to load data first" -ForegroundColor Red
    break
}
$dataToProcess = $bulkData
$dataToProcess = $dataToProcess | Where-Object {$_."OA Import Status" -eq "READY FOR IMPORT"}
$decision = Read-Host "You're about to create $($dataToProcess.Count) accounts. Are you sure? (type yes)"
if($decision.ToLower() -ne "yes"){
    Write-Host "User account creation aborted"
    break
}
$importList = @()
foreach($row in $dataToProcess){
    $userObj = @{}
    $userObj.firstName = $row."First Name"
    $userObj.lastName = $row."Last Name"
    $userObj.userEmail = $row.Email
    $userObj.parameters = @{}
    $userObj.parameters.nickname = $row."User ID"
    $userObj.parameters.line_managerid = $row.Manager
    $userObj.parameters.departmentid = $row.Department
    $userObj.parameters.job_codeid = $row."Job code"
    $userObj.parameters.UserCountry__c = $row."User Country"
    $userObj.parameters.EmploymentStatus__c = $row."Employment status"
    $userObj.parameters.Contract_type__c = $row."Contract type"
    $userObj.parameters.JobFunction__c = $row."Functions For Utilisation"
    $userObj.parameters.Company__c = $row.Company
    $userObj.parameters.UserLocation__c = $row.Location
    $userObj.parameters.CoE__c = $row.CoE
    $userObj.parameters.Clan__c = $row.Clan
    $userObj.parameters.Billability__c = $row.Billability
    $userObj.parameters.VaultCode__c = $row.VaultCode
    $userObj.parameters.active = ($row."Is Active" -eq "Active") ? 1 : 0
    $userObj.parameters.rate = $row.Cost
    $userObj.parameters.password = $row.Password
    $importList += $userObj
}
$counter = [pscustomobject] @{ Value = 0 }
[int]$groupSize = Read-Host "Please provide current license limit"
$groups = $importList | Group-Object -Property { [math]::Floor($counter.Value++ / $groupSize) }
if($groupSize -lt $importList.Count){
    Write-Host "List of users exceeds license limit for one request. Users are divided to $($groups.Count) groups"
}
foreach($group in $groups){
    $params = @{}
    $params.usersData = $group.Group
    $resp = $connector.SendRequest([OARequestType]::CreateUserBulk,$params)
    Write-Host "Saving transaction file"
    $transactionID = New-Guid
    $successIDs = ($resp.response.CreateUser.User | Select-Object -Property id).id
    $createdUserIDs = ($resp.response.CreateUser.User | Select-Object -Property nickname).nickname
    Set-Content -Path "$logFolder/createdUsers-$transactionID.txt" ($createdUserIDs -join ';')
    Set-Content -Path "$logFolder/$transactionID.txt" ($successIDs -join ';')
    $failedRequests = @()
    if ($resp.response.CreateUser.Count -eq 1){
        if($resp.response.CreateUser.status -ne 0){
            $errorResp = $connector.SendRequest([OARequestType]::Read,@{limit="1";type="Error";method="equal to";queryData=@{code=($resp.response.CreateUser[$i]).status}})
            $failedRequests += @{
                First=$params.usersData[$i].firstName;
                Last=$params.usersData[$i].lastName;
                "Error code"=($resp.response.CreateUser[$i]).status;
                "Error text"=$errorResp.response.Read.Error.text;
                "OuterXml" = $resp.response.CreateUser[$1].OuterXml
            }
        }
    }
    else {
        for($i=0;$i -lt $resp.response.CreateUser.Count; $i++){
            if(($resp.response.CreateUser[$i]).status -ne 0){
                $errorResp = $connector.SendRequest([OARequestType]::Read,@{limit="1";type="Error";method="equal to";queryData=@{code=($resp.response.CreateUser[$i]).status}})
                $failedRequests += @{
                    First=$params.usersData[$i].firstName;
                    Last=$params.usersData[$i].lastName;
                    "Error code"=($resp.response.CreateUser[$i]).status;
                    "Error text"=$errorResp.response.Read.Error.text
                }
            }
        }
    }
    Set-Content -Path "$logFolder/error-$transactionID.txt" ($failedRequests | Format-List | Out-String)
    Write-Host "Out of $($params.usersData.Count):`n`t$($successIDs.Count) were created successfully`n`t$($failedRequests.Count) failed"
    Write-Host "Transaction ID: $transactionID"
    $usersToDisable = (Get-Content -Path "$logFolder/$transactionID.txt") -split ";"
    $modRequests = @()
    foreach($id in $usersToDisable){
        $modRequest = @{
            type = "User";
            id = $id;
            dataToUpdate = @{
                active = "0"
            }
        }
        $modRequests += $modRequest
    }
    $resp = $connector.SendRequest([OARequestType]::ModifyBulk,@{modifyRequests = $modRequests})
    Write-Host "$($resp.response.Modify.User.Count) users were deactivated"
}