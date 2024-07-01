if(-not $usersToCreate){
    Write-Host "You need to load data first" -ForegroundColor Red
    break
}

$OAStatusReady = "READY FOR IMPORT"
$OACleaningStatusValidate = "NEEDS VALIDATION"
$OAStatusImported = "IMPORTED"
$OAStatusSkip = "IMPORTED"
$OARegularAccount = "Regular"
$OASkeletonAccount = "Skeleton"

$dataToProcess = $usersToCreate
$dataToProcess = $dataToProcess | Where-Object {$_."OA Import Status" -ne $OAStatusSkip}
$decision = Read-Host "You're about to create $($dataToProcess.Count) accounts. Are you sure? (type yes)"
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
    $userObj = @{}
    $userObj.firstName = $row."First Name"
    $userObj.lastName = $row."Last Name"
    $userObj.userEmail = $row.Email
    $userObj.parameters = @{}
    $userObj.parameters.nickname = $row."User ID"
    $userObj.parameters.Company__c = $row.Company
    $userObj.parameters.UserCountry__c = $row."User Country"
    $userObj.parameters.UserLocation__c = $row.Location
    $userObj.parameters.line_managerid = $userDataDictionary[$row.Manager] ?? "1"
    $userObj.parameters.JobTitle__c = $row."Job Title"
    $userObj.parameters.CoE__c = $row.CoE
    $userObj.parameters.departmentid = $departmentDataDictionary[$row.Department] ?? "48"
    $userObj.parameters.Stream__c = $row.Stream
    $userObj.parameters.Clan__c = $row.Clan -eq "<empty>" ? "" : $row.Clan
    $userObj.parameters.job_codeid = $jobcodeDataDict[$row."Job code"] ?? "94"
    $userObj.parameters.VaultCode__c = $row.VaultCode
    $userObj.parameters.EmploymentStatus__c = $row."Employment status"
    $userObj.parameters.JobFunction__c = $row."Functions For Utilisation"
    $userObj.parameters.Contract_type__c = $row."Contract type"
    $userObj.parameters.Billability__c = $row.Billability -eq "<empty>" ? "" : $row.Billability
    if($row."Timesheets are approved by" -eq "Billable Time Only Approval Process") {
        $userObj.parameters.ta_approvalprocess = "11"
        $userObj.parameters.ta_approver = ""
    }else{
        $userObj.parameters.ta_approver = "-1"
    }
    switch($row.Schedule){
        "UK work schedule - 35h" {
            $userObj.parameters.account_workscheduleid = "2484"
        }
        "UK work schedule - 37,5h" {
            $userObj.parameters.account_workscheduleid = "2482"
        }
        "UK work schedule - 40h" {
            $userObj.parameters.account_workscheduleid = "2481"
        }
        "Portugal work schedule - 37,5h" {
            $userObj.parameters.account_workscheduleid = "2483"
        }
        "China work schedule - 37,5h" {
            $userObj.parameters.account_workscheduleid = "2480"
        }
        "<set manually>" {
            switch($row."User Country"){
                "UK" {
                    $userObj.parameters.account_workscheduleid = "2482"
                }
                "PT" {
                    $userObj.parameters.account_workscheduleid = "2483"
                }
                "IN" {
                    $userObj.parameters.account_workscheduleid = "2482"
                }
                "CN" {
                    $userObj.parameters.account_workscheduleid = "2480"
                }
                "NL" {
                    $userObj.parameters.account_workscheduleid = "2482"
                }
                "HU" {
                    $userObj.parameters.account_workscheduleid = "2481"
                }
                "SA" {
                    $userObj.parameters.account_workscheduleid = "2481"
                }
                "US" {
                    $userObj.parameters.account_workscheduleid = "2481"
                }
            }
        }
        "" {
            switch($row."User Country"){
                "UK" {
                    $userObj.parameters.account_workscheduleid = "2482"
                }
                "PT" {
                    $userObj.parameters.account_workscheduleid = "2483"
                }
                "IN" {
                    $userObj.parameters.account_workscheduleid = "2482"
                }
                "CN" {
                    $userObj.parameters.account_workscheduleid = "2480"
                }
                "NL" {
                    $userObj.parameters.account_workscheduleid = "2482"
                }
                "HU" {
                    $userObj.parameters.account_workscheduleid = "2481"
                }
                "SA" {
                    $userObj.parameters.account_workscheduleid = "2481"
                }
                "US" {
                    $userObj.parameters.account_workscheduleid = "2481"
                }
            }
        }
    }
    switch($row.Role){
        ""{
            $userObj.parameters.role_id = "2"
            $userObj.parameters.primary_filter_set = "2"
        }
        "Administrator" {
            $userObj.parameters.role_id = "1"
            $userObj.parameters.primary_filter_set = "1"
        }
        "User" {
            $userObj.parameters.role_id = "2"
            $userObj.parameters.primary_filter_set = "2"
        }
        "Solution Delivery Consultant" {
            $userObj.parameters.role_id = "3"
            $userObj.parameters.primary_filter_set = "1"
        }
        "old_Finance / Admin PL" {
            $userObj.parameters.role_id = "5"
            $userObj.parameters.primary_filter_set = "1"
        }
        "Line Manager" {
            $userObj.parameters.role_id = "7"
            $userObj.parameters.primary_filter_set = "1"
        }
        "old_Finance / Admin UK" {
            $userObj.parameters.role_id = "8"
            $userObj.parameters.primary_filter_set = "1"
        }
        "Practice Head" {
            $userObj.parameters.role_id = "9"
            $userObj.parameters.primary_filter_set = "1"
        }
        "Head of Engineering" {
            $userObj.parameters.role_id = "10"
            $userObj.parameters.primary_filter_set = "1"
        }
        "Technical Project Manager" {
            $userObj.parameters.role_id = "11"
            $userObj.parameters.primary_filter_set = "1"
        }
        "Sales and Marketing support" {
            $userObj.parameters.role_id = "12"
            $userObj.parameters.primary_filter_set = "1"
        }
        "User + Reports" {
            $userObj.parameters.role_id = "13"
            $userObj.parameters.primary_filter_set = "2"
        }
        "Junior IT Administrator" {
            $userObj.parameters.role_id = "14"
            $userObj.parameters.primary_filter_set = "1"
        }
        "OA Extension" {
            $userObj.parameters.role_id = "15"
            $userObj.parameters.primary_filter_set = "1"
        }
        "New role" {
            $userObj.parameters.role_id = "16"
            $userObj.parameters.primary_filter_set = "1"
        }
        "Finance User + Reports" {
            $userObj.parameters.role_id = "17"
            $userObj.parameters.primary_filter_set = "1"
        }
        "Technology Director" {
            $userObj.parameters.role_id = "18"
            $userObj.parameters.primary_filter_set = "1"
        }
        "User + Timesheets" {
            $userObj.parameters.role_id = "19"
            $userObj.parameters.primary_filter_set = "1"
        }
        "UK HR" {
            $userObj.parameters.role_id = "20"
            $userObj.parameters.primary_filter_set = "1"
        }
        "US HR" {
            $userObj.parameters.role_id = "21"
            $userObj.parameters.primary_filter_set = "1"
        }
        "JML Process Automation" {
            $userObj.parameters.role_id = "22"
            $userObj.parameters.primary_filter_set = "1"
        }
        "Client Services" {
            $userObj.parameters.role_id = "23"
            $userObj.parameters.primary_filter_set = "1"
        }
        "User + user administrator" {
            $userObj.parameters.role_id = "24"
            $userObj.parameters.primary_filter_set = "1"
        }
        "old_Finance / Admin PL + Invoices" {
            $userObj.parameters.role_id = "25"
            $userObj.parameters.primary_filter_set = "1"
        }
        "Finance / Manager PL" {
            $userObj.parameters.role_id = "27"
            $userObj.parameters.primary_filter_set = "1"
        }
        "Full View" {
            $userObj.parameters.role_id = "28"
            $userObj.parameters.primary_filter_set = "1"
        }
        "Finance / Manager PL + Invoices" {
            $userObj.parameters.role_id = "29"
            $userObj.parameters.primary_filter_set = "1"
        }
        "User + User Manager" {
            $userObj.parameters.role_id = "30"
            $userObj.parameters.primary_filter_set = "1"
        }
        "Finance / Manager UK" {
            $userObj.parameters.role_id = "31"
            $userObj.parameters.primary_filter_set = "1"
        }
        "Finance / Manager UK + Invoices" {
            $userObj.parameters.role_id = "32"
            $userObj.parameters.primary_filter_set = "1"
        }
        "Project Accountant" {
            $userObj.parameters.role_id = "33"
            $userObj.parameters.primary_filter_set = "1"
        }
    }
    $NewPassword = ("ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz".ToCharArray() | Sort-Object { Get-Random })[0..14] -join ''
    $NewPassword += ("0123456789".ToCharArray() | Sort-Object { Get-Random })[0..2] -join ''
    $NewPassword += ("~\\!@#$%^&*()-_=+[]{}|;:,.<>/?".ToCharArray() | Sort-Object { Get-Random })[0] -join ''
    $userObj.parameters.password = $NewPassword
    $userObj.parameters.active = ($row."Is Active" -eq "Yes") ? 1 : 0
    $userObj.parameters.saml_auth__c = 0
    $userObj.parameters.ExternalCostCentre__c = "1"
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
            $errorResp = $connector.SendRequest([OARequestType]::Read,@{limit="1";type="Error";method="equal to";queryData=@{code=($resp.response.CreateUser[$i]).status}})
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