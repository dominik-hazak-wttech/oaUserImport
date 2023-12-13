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
    $connector = [OAConnector]::new($config.namespace, $config.apiKey)
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
                    continue
                }
                elseif ($prompt.toLower() -eq "n"){
                    $credLoop = $false
                }
                else{
                    Write-Host "Please type 'y' or 'n'"
                }
            }
            while($credLoop)
        }
        return 1
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
    Write-Host "3. Create user in OpenAir"
    Write-Host "0. Exit"
    $prompt = Read-Host "Your choice [0-3]"

    switch($prompt){
        1 {
            Write-Host "This feature is not yet implemented"
            # Write-Host "Please provide path for data file"
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
        0{
            $looping = $false
        }
        default{
            continue
        }
    }
}while($looping)
Write-Host "Have a nice day :)"