Import-Module PSExcel
. ./OAConnector.ps1

$logFolder = "./logs"
if(-not (Test-Path -Path $logFolder)){
    New-Item -Path $logFolder -ItemType Directory
}
Write-Host "Loading configuration"
$config = Get-Content -Path ./config.json | ConvertFrom-Json 
Write-Host "Setting up connection to OpenAir"
$connected = $false
do{
    if(-not $config){
        $defaultConfig = '
        {
            "instances":[
                {
                    "name":"default",
                    "namespace":"default",
                    "apiKey":"",
                    "company":"",
                    "login":""
                }
            ]
        }
        '
        Set-Content -Path ./config.json $defaultConfig
        Write-Host "Config file not found. Default was created please fill the data"
        return 1
    }
    Write-Host "Available instances"
    for($i = 0; $i -lt $config.instances.Count;$i++){
        Write-Host "`t$i - $($config.instances[$i].name)"
    }
    [int]$instance = Read-Host "Select number of instance"
    if(-not ($config.instances[$instance].company -or $config.instances[$instance].login)){
    $connector = [OAConnector]::new($config.instances[$instance].namespace, $config.instances[$instance].apiKey, $config.instances[$instance].url)
    }
    else{
        $connector = [OAConnector]::new($config.instances[$instance].namespace, $config.instances[$instance].apiKey, $config.instances[$instance].url, $config.instances[$instance].company, $config.instances[$instance].login)
    }
    Write-Host "Testing connection to OpenAir API...  " -NoNewline
    $timeResp = $connector.SendRequest([OARequestType]::Time)
    $timeOK = $timeResp.response.Time.status -eq 0
    if($timeOK){
        Write-Host "success"
    }
    else{
        Write-Host "failed"
        Write-Host "Connection to OpenAir api cannot be made. Please check internet connection and try again."
        return 1
    }
    Write-Host "Authorizing in OpenAir API...  " -NoNewline
    $timeResp = $connector.SendRequest([OARequestType]::Auth)
    $authOK = $timeResp.response.Auth.status -eq 0
    if($authOK){
        Write-Host "success"
    }
    else{
        Write-Host "failed"
        if($timeResp.response.Auth.status -eq 401){
            Write-Host "Provided credentials are wrong."
            $credLoop = $true
            do{
                $prompt = Read-Host -Prompt "Do you want to provide them again? (Y/N)"
                if ($prompt.toLower() -eq "y"){
                    break
                }
                elseif ($prompt.toLower() -eq "n"){
                    return 1
                }
                else{
                    Write-Host "Please type 'y' or 'n'"
                }
            }
            while($credLoop)
        }
    }
    $connected = $authOK -and $timeOK
}
while(-not $connected)
Write-Host "OpenAir API ready. Please select action to perform"
$looping = $true
do{
    Write-Host "Menu:"
    Write-Host "1. Read data from file"
    Write-Host "2. Read data from OpenAir"
    Write-Host "3. Create user in OpenAir (single)"
    if($bulkData){
        Write-Host "4. Create user in OpenAir (bulk from data)"
        Write-Host "5. Generate request for bulk user creation"
        Write-Host "6. Add all users within the license limit "
    }
    else{
        Write-Host "4. Create user in OpenAir (bulk from data)" -ForegroundColor DarkGray
        Write-Host "5. Generate request for bulk user creation" -ForegroundColor DarkGray
        Write-Host "6. Add all users within the license limit " -ForegroundColor DarkGray
    }
    Write-Host "7. Revert user creation"
    Write-Host "0. Exit"
    $prompt = Read-Host "Your choice [0-7]"

    switch($prompt){
        1 {
            $path = Read-Host "Please provide path for data file"
            $type = Read-Host "Please provide file type (csv, xls)"
            switch($type.ToLower()){
                xls {
                    $sheet = Read-Host "Please provide Sheet name (with no name default sheet will be selected)"
                    $rowNum = Read-Host "Please starting row number (by default first row will be chosen)"
                    if ($sheet -and $rowNum){
                        $bulkData = Import-XLSX -Path $path -Sheet $sheet -RowStart $rowNum
                    }
                    elseif ($rowNum){
                        $bulkData = Import-XLSX -Path $path -RowStart $rowNum
                    }
                    elseif ($sheet){
                        $bulkData = Import-XLSX -Path $path -Sheet $sheet
                    }
                    else{
                        $bulkData = Import-XLSX -Path $path
                    }
                }
                csv {
                    $delim = Read-Host "What character divides data?"
                    $bulkData = Import-Csv -Path $path -Delimiter $delim
                }
                default{
                    Write-Host "Unknown file type $type. Please try again"
                }
            }
        }
        2{
            $type = Read-Host "Provide OpenAir Object Type to search for"
            $method = Read-Host "Provide method of the Read action"
            $queryData = Read-Host "Provide read request data in form of 'key=value,key=value(...)'"
            $limit = Read-Host "Provide limit of response entries (default=10)"
            $enableCustom = Read-Host "Emable custom fields (default=false)"

            $queryDataSplit = $queryData -split ","
            $queryDataTable = @{}
            foreach($row in $queryDataSplit){
                $rowSplit = $row.Trim() -split "="
                $queryDataTable[$rowSplit[0]] = $rowSplit[1]
            }

            $params = @{type=$type;method=$method;queryData=$queryDataTable}
            if ($limit){
                $params.limit = $limit
            }
            if ($enableCustom){
                $params.customFields = $false
                if($enableCustom -eq "true" -or $enableCustom -eq "1" -or $enableCustom -eq "yes"){
                    $params.customFields = $true
                }
            }
            $resp = $connector.SendRequest([OARequestType]::Read,$params)
            Write-Host ($resp.response.Read.User | Format-List | Out-String)
        }
        3 {
            $params = @{}
            $params.firstName = Read-Host "Provide user's first name"
            $params.lastName = Read-Host "Provide user's last name"
            $params.userEmail = Read-Host "Provide user's email"
            
            $parameters = @{}
            $parameters.nickname = Read-Host "Provide data for username"
            $parameters.line_managerid = Read-Host "Provide data for line manager (id)"
            $parameters.departmentid = Read-Host "Provide data for department (id)"
            $parameters.job_codeid = Read-Host "Provide data for job code (id)"
            $parameters.UserCountry__c = Read-Host "Provide data for user country"
            $parameters.EmploymentStatus__c = Read-Host "Provide data for employment status"
            $parameters.Contract_type__c = Read-Host "Provide data for contract type"
            $parameters.JobFunction__c = Read-Host "Provide data for functions for utilisation"
            $parameters.Company__c = Read-Host "Provide data for company"
            $parameters.UserLocation__c = Read-Host "Provide data for location"
            $parameters.CoE__c = Read-Host "Provide data for CoE"
            $parameters.Clan__c = ""
            $parameters.Billability__c = Read-Host "Provide data for billability"
            $parameters.VaultCode__c = ""
            $parameters.active = "1"
            $parameters.rate = ""
            $parameters.password = Read-Host "Provide password for user"
            $params.parameters = $parameters
            $resp = $connector.SendRequest([OARequestType]::CreateUser,$params)
            Write-Host ($resp.response.CreateUser.User | Format-List | Out-String)
        }
        4{
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
            $groupSize = 999
            $groups = $importList | Group-Object -Property { [math]::Floor($counter.Value++ / $groupSize) }
            if($groupSize -lt $importList.Count){
                Write-Host "List of users exceeds API limit for one request. Users are divided to $($groups.Count) groups"
            }
            foreach($group in $groups){
                $params = @{}
                $params.usersData = $group.Group
                $resp = $connector.SendRequest([OARequestType]::CreateUserBulk,$params)
                Write-Host "Saving transaction file"
                $transactionID = New-Guid
                $successIDs = ($resp.response.CreateUser | Where-Object {$_.status -eq "0"} | Select-Object -Property id).id
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
                else{
                    for($i=0;$i -lt $resp.response.CreateUser.Count; $i++){
                        if(($resp.response.CreateUser[$i]).status -ne 0){
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
                }
                Set-Content -Path "$logFolder/error-$transactionID.txt" ($failedRequests | Format-List | Out-String)
                Write-Host "Out of $($params.usersData.Count):`n`t$($successIDs.Count) were created successfully`n`t$($failedRequests.Count) failed"
                Write-Host "Transaction ID: $transactionID"
                Set-Content -Path "$logFolder/response-$transactionID.xml" $resp.response.OuterXml
            }
        }
        5{
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
            $groupSize = 999
            $groups = $importList | Group-Object -Property { [math]::Floor($counter.Value++ / $groupSize) }
            if($groupSize -lt $importList.Count){
                Write-Host "List of users exceeds API limit for one request. Users are divided to $($groups.Count) groups"
            }
            foreach($group in $groups){
                $request = $connector.GenerateCreateUserBulkRequest($group.Group)
                Write-Host ($request.OuterXml)
            }
        }
        6{
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
                Set-Content -Path "$logFolder/$transactionID.txt" ($successIDs -join ';')
                $failedRequests = @()
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
        }
        7 {
            $transactionID = Read-Host "Provide Transaction ID"
            $idsToRemove = ((Get-Content -Path "$logFolder/$transactionID.txt") -split ';')
            Write-Host "IDs to remove: $idsToRemove"
            $resp = $connector.SendRequest([OARequestType]::DeleteUser,@{userIDs=$idsToRemove})
            $resp.response.OuterXml
        }
        8 {
            $dataLeft = $true
            $i = 0
            $activeUsers = @()
            do{
                $resp = $connector.SendRequest([OARequestType]::Read,@{type="User";method="not equal to";queryData=@{role_id="1"};limit="$i,1000"})

                if ($resp.response.Read.ChildNodes.Count -ne 1000){
                    $dataLeft = $false 
                }
                $activeUsers += ($resp.response.Read.User | Where-Object {$_.active})
                Write-Host ($resp.response.Read.User | Where-Object {$_.active}).id
                $i += 1000
            } while($dataLeft)
            Set-Content -Path "$logFolder/UsersToDisable.txt" (($activeUsers.id | Get-Unique)-join';')
        }
        9{
            $usersToDisable = (Get-Content -Path "$logFolder/UsersToDisable.txt") -split ";"
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
            Write-Host $resp.response.OuterXml
        }
        0{
            $looping = $false
        }
        default{
            continue
        }
    }
}while($looping)
Write-Host "Have a nice day :)"