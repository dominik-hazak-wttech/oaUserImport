Import-Module PSExcel
. ./OAConnector.ps1

Write-Host "Loading configuration"
$config = Get-Content -Path ./config.json | ConvertFrom-Json 
Write-Host "Setting up connection to OpenAir"
$connected = $false
do{
    if(-not $config){
        Write-Host "Config file not found. Please provide API connection data"
        $config = @{}
        $config.namespace = Read-Host -Prompt "Namespace"
        $config.apiKey = Read-Host -Prompt "API key"
    }
    if(-not ($config.company -or $config.login)){
    $connector = [OAConnector]::new($config.namespace, $config.apiKey)
    }
    else{
        $connector = [OAConnector]::new($config.namespace, $config.apiKey, $config.company, $config.login)
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
    }
    else{
        Write-Host "4. Create user in OpenAir (bulk from data)" -ForegroundColor DarkGray
    }
    Write-Host "0. Exit"
    $prompt = Read-Host "Your choice [0-4]"

    switch($prompt){
        1 {
            $path = Read-Host "Please provide path for data file"
            $type = Read-Host "Please provide file type (csv, xls)"
            switch($type.ToLower()){
                xls {
                    $bulkData = Import-XLSX -Path $path
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
            $params.parameters = $parameters
            $resp = $connector.SendRequest([OARequestType]::CreateUser,$params)
            Write-Host ($resp.response.CreateUser.User | Format-List | Out-String)
        }
        4{
            if(-not $bulkData){
                Write-Host "You need to load data first" -ForegroundColor Red
                break
            }
            Write-Host "Not yet implemented"
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