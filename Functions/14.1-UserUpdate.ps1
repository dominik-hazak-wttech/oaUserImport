if(-not $usersToUpdate){
    Write-Host "You need to load data first" -ForegroundColor Red
    break
}

$OAStatusReady = "READY FOR IMPORT"
$OACleaningStatusValidate = "NEEDS VALIDATION"
$OAStatusImported = "IMPORTED"
$OAStatusSkip = "SKIP"
$OARegularAccount = "Regular"
$OASkeletonAccount = "Skeleton"

$dataToProcess = $usersToUpdate
$dataToProcess = $dataToProcess | Where-Object {$_."OA Import Status" -ne $OAStatusSkip}
$decision = Read-Host "You're about to udpate $($dataToProcess.Count) accounts. Are you sure? (type yes)"
if($decision.ToLower() -ne "yes"){
    Write-Host "User account creation aborted"
    break
}
$importList = @()
foreach($row in $dataToProcess){
    if (
        $row."First Name" -eq "" -or 
        $null -eq $row."First Name" -or
        $row."Last Name" -eq "" -or 
        $null -eq $row."Last Name"
       ){
        Write-Error "$($row.Email) is missing First or Last name!"
        $failState = $true
    }
    if ($row.Email -eq "" -or $null -eq $row.Email){
        Write-Error "$($row."First Name") $($row."Last Name") is missing email value!"
        $failState = $true
    }
    if ($row."User ID" -eq "" -or $null -eq $row."User ID"){
        Write-Error "$($row.Email) is missing User ID value!"
        $failState = $true
    }
    if (-not $userDataDictionary[$row."User ID"]){
        Write-Warning "$($row."User ID") is missing id in dictionary! User will be skipped!"
        continue
    }
    $userObj = @{}
    $userObj.id = $userDataDictionary[$row."User ID"]
    $userObj.type = "User"
    $userObj.dataToUpdate = @{}
    $userObj.dataToUpdate.firstName = $row."First Name"
    $userObj.dataToUpdate.lastName = $row."Last Name"
    $userObj.dataToUpdate.userEmail = $row.Email
    $userObj.dataToUpdate.Company__c = $row.Company
    $userObj.dataToUpdate.UserCountry__c = $row."User Country"
    $userObj.dataToUpdate.nickname = $row."User ID"
    $userObj.dataToUpdate.UserLocation__c = $row.Location
    $userObj.dataToUpdate.line_managerid = $userDataDictionary[$row.Manager] ?? "1"
    $userObj.dataToUpdate.JobTitle__c = $row."Job Title"
    $userObj.dataToUpdate.CoE__c = $row.CoE
    $userObj.dataToUpdate.departmentid = $departmentDataDictionary[$row.Department] ?? "48"
    $userObj.dataToUpdate.Stream__c = $row.Stream
    $userObj.dataToUpdate.Clan__c = ($row.Clan -eq "<empty>") ? "" : $row.Clan
    $userObj.dataToUpdate.Contract_type__c = $row."Contract type"
    if($row."Timesheets are approved by" -eq "Billable Time Only Approval Process") {
        $userObj.dataToUpdate.ta_approvalprocess = "11"
        $userObj.dataToUpdate.ta_approver = ""
    }else{
        $userObj.dataToUpdate.ta_approver = "-1"
    }
    switch($row.Schedule){
        "UK work schedule - 35h" {
            $userObj.dataToUpdate.account_workscheduleid = "2484"
        }
        "UK work schedule - 37,5h" {
            $userObj.dataToUpdate.account_workscheduleid = "2482"
        }
        "UK work schedule - 40h" {
            $userObj.dataToUpdate.account_workscheduleid = "2481"
        }
        "Portugal work schedule - 37,5h" {
            $userObj.dataToUpdate.account_workscheduleid = "2483"
        }
        "China work schedule - 37,5h" {
            $userObj.dataToUpdate.account_workscheduleid = "2480"
        }
        "<set manually>" {
            switch($row."User Country"){
                "UK" {
                    $userObj.dataToUpdate.account_workscheduleid = "2482"
                }
                "PT" {
                    $userObj.dataToUpdate.account_workscheduleid = "2483"
                }
                "IN" {
                    $userObj.dataToUpdate.account_workscheduleid = "2482"
                }
                "CN" {
                    $userObj.dataToUpdate.account_workscheduleid = "2480"
                }
                "NL" {
                    $userObj.dataToUpdate.account_workscheduleid = "2482"
                }
                "HU" {
                    $userObj.dataToUpdate.account_workscheduleid = "2481"
                }
                "SA" {
                    $userObj.dataToUpdate.account_workscheduleid = "2481"
                }
                "US" {
                    $userObj.dataToUpdate.account_workscheduleid = "2481"
                }
            }
        }
        "" {
            switch($row."User Country"){
                "UK" {
                    $userObj.dataToUpdate.account_workscheduleid = "2482"
                }
                "PT" {
                    $userObj.dataToUpdate.account_workscheduleid = "2483"
                }
                "IN" {
                    $userObj.dataToUpdate.account_workscheduleid = "2482"
                }
                "CN" {
                    $userObj.dataToUpdate.account_workscheduleid = "2480"
                }
                "NL" {
                    $userObj.dataToUpdate.account_workscheduleid = "2482"
                }
                "HU" {
                    $userObj.dataToUpdate.account_workscheduleid = "2481"
                }
                "SA" {
                    $userObj.dataToUpdate.account_workscheduleid = "2481"
                }
                "US" {
                    $userObj.dataToUpdate.account_workscheduleid = "2481"
                }
            }
        }
    }
    switch($row.Role){
        ""{
            $userObj.dataToUpdate.role_id = "2"
            $userObj.dataToUpdate.primary_filter_set = "2"
        }
        "Administrator" {
            $userObj.dataToUpdate.role_id = "1"
            $userObj.dataToUpdate.primary_filter_set = "1"
        }
        "User" {
            $userObj.dataToUpdate.role_id = "2"
            $userObj.dataToUpdate.primary_filter_set = "2"
        }
        "Solution Delivery Consultant" {
            $userObj.dataToUpdate.role_id = "3"
            $userObj.dataToUpdate.primary_filter_set = "1"
        }
        "old_Finance / Admin PL" {
            $userObj.dataToUpdate.role_id = "5"
            $userObj.dataToUpdate.primary_filter_set = "1"
        }
        "Line Manager" {
            $userObj.dataToUpdate.role_id = "7"
            $userObj.dataToUpdate.primary_filter_set = "1"
        }
        "old_Finance / Admin UK" {
            $userObj.dataToUpdate.role_id = "8"
            $userObj.dataToUpdate.primary_filter_set = "1"
        }
        "Practice Head" {
            $userObj.dataToUpdate.role_id = "9"
            $userObj.dataToUpdate.primary_filter_set = "1"
        }
        "Head of Engineering" {
            $userObj.dataToUpdate.role_id = "10"
            $userObj.dataToUpdate.primary_filter_set = "1"
        }
        "Technical Project Manager" {
            $userObj.dataToUpdate.role_id = "11"
            $userObj.dataToUpdate.primary_filter_set = "1"
        }
        "Sales and Marketing support" {
            $userObj.dataToUpdate.role_id = "12"
            $userObj.dataToUpdate.primary_filter_set = "1"
        }
        "User + Reports" {
            $userObj.dataToUpdate.role_id = "13"
            $userObj.dataToUpdate.primary_filter_set = "2"
        }
        "Junior IT Administrator" {
            $userObj.dataToUpdate.role_id = "14"
            $userObj.dataToUpdate.primary_filter_set = "1"
        }
        "OA Extension" {
            $userObj.dataToUpdate.role_id = "15"
            $userObj.dataToUpdate.primary_filter_set = "1"
        }
        "New role" {
            $userObj.dataToUpdate.role_id = "16"
            $userObj.dataToUpdate.primary_filter_set = "1"
        }
        "Finance User + Reports" {
            $userObj.dataToUpdate.role_id = "17"
            $userObj.dataToUpdate.primary_filter_set = "1"
        }
        "Technology Director" {
            $userObj.dataToUpdate.role_id = "18"
            $userObj.dataToUpdate.primary_filter_set = "1"
        }
        "User + Timesheets" {
            $userObj.dataToUpdate.role_id = "19"
            $userObj.dataToUpdate.primary_filter_set = "1"
        }
        "UK HR" {
            $userObj.dataToUpdate.role_id = "20"
            $userObj.dataToUpdate.primary_filter_set = "1"
        }
        "US HR" {
            $userObj.dataToUpdate.role_id = "21"
            $userObj.dataToUpdate.primary_filter_set = "1"
        }
        "JML Process Automation" {
            $userObj.dataToUpdate.role_id = "22"
            $userObj.dataToUpdate.primary_filter_set = "1"
        }
        "Client Services" {
            $userObj.dataToUpdate.role_id = "23"
            $userObj.dataToUpdate.primary_filter_set = "1"
        }
        "User + user administrator" {
            $userObj.dataToUpdate.role_id = "24"
            $userObj.dataToUpdate.primary_filter_set = "1"
        }
        "old_Finance / Admin PL + Invoices" {
            $userObj.dataToUpdate.role_id = "25"
            $userObj.dataToUpdate.primary_filter_set = "1"
        }
        "Finance / Manager PL" {
            $userObj.dataToUpdate.role_id = "27"
            $userObj.dataToUpdate.primary_filter_set = "1"
        }
        "Full View" {
            $userObj.dataToUpdate.role_id = "28"
            $userObj.dataToUpdate.primary_filter_set = "1"
        }
        "Finance / Manager PL + Invoices" {
            $userObj.dataToUpdate.role_id = "29"
            $userObj.dataToUpdate.primary_filter_set = "1"
        }
        "User + User Manager" {
            $userObj.dataToUpdate.role_id = "30"
            $userObj.dataToUpdate.primary_filter_set = "1"
        }
        "Finance / Manager UK" {
            $userObj.dataToUpdate.role_id = "31"
            $userObj.dataToUpdate.primary_filter_set = "1"
        }
        "Finance / Manager UK + Invoices" {
            $userObj.dataToUpdate.role_id = "32"
            $userObj.dataToUpdate.primary_filter_set = "1"
        }
        "Project Accountant" {
            $userObj.dataToUpdate.role_id = "33"
            $userObj.dataToUpdate.primary_filter_set = "1"
        }
    }
    $userObj.dataToUpdate.active = ($row."Is Active" -eq "Yes") ? "1" : "0"
    $userObj.dataToUpdate.saml_auth__c = "0"
    $userObj.dataToUpdate.ExternalCostCentre__c = "1"
    if($row."OA user TYPE" -eq $OASkeletonAccount){
        $userObj.dataToUpdate.job_codeid = $jobcodeDataDict[$row."Job code"] ?? "94"
        $userObj.dataToUpdate.EmploymentStatus__c = $row."Employment status"
        $userObj.dataToUpdate.JobFunction__c = $row."Functions For Utilisation"
        $userObj.dataToUpdate.Billability__c = ($row.Billability -eq "<empty>") ? "" : $row.Billability
    }
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
    $params.modifyRequests = $group.Group
    if($validateOnly){
        Write-Host "Would send request to edit group of $($group.Group.Count) accounts"
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