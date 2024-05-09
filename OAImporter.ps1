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
    if($config.instances[$instance].access_token){
        $connector = [OAConnector]::new($config.instances[$instance].namespace, $config.instances[$instance].apiKey, $config.instances[$instance].url, $config.instances[$instance].access_token)
    }
    elseif($config.instances[$instance].company -and $config.instances[$instance].login){
        $connector = [OAConnector]::new($config.instances[$instance].namespace, $config.instances[$instance].apiKey, $config.instances[$instance].url, $config.instances[$instance].company, $config.instances[$instance].login)
    }
    else{
        $connector = [OAConnector]::new($config.instances[$instance].namespace, $config.instances[$instance].apiKey, $config.instances[$instance].url)
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
    Write-Host "8. Get all users except admins (and save in UsersToDisable.txt)"
    Write-Host "9. Deactivate selected users (from UsersToDisable.txt)"
    Write-Host "10. Activate selected users (from UsersToDisable.txt)"
    if($bulkData){
        Write-Host "11. Migrate login to SSO"
    }
    else{
        Write-Host "11. Migrate login to SSO" -ForegroundColor DarkGray
    }
    Write-Host "0. Exit"
    $prompt = Read-Host "Your choice [0-7]"

    switch($prompt){
        1 {
            . ./Functions/1-ReadDataFromFile.ps1
        }
        2 {
            . ./Functions/2-ReadDataFromOpenAir.ps1
        }
        3 {
            . ./Functions/3-CreateSingleUserInOpenAir.ps1
        }
        4 {
            . ./Functions/4-CreateBulkUsersInOpenAir.ps1
        }
        5 {
            . ./Functions/5-GenerateRequestForBulkCreation.ps1
        }
        6 {
            . ./Functions/6-AddAllUsersWithinTheLicenseLimit.ps1
        }
        7 {
            . ./Functions/7-RevertUserCreation.ps1
        }
        #additional things for test
        8 {
            . ./Functions/8-GetAllUsersExceptAdmins.ps1
        }
        9 {
            . ./Functions/9-DeactivateSelectedUsers.ps1
        }
        10 {
            . ./Functions/10-ActivateSelectedUsers.ps1
        }
        11 {
            . ./Functions/11-MigrateLoginToSSO.ps1
        }
        0 {
            $looping = $false
        }
        default{
            continue
        }
    }
}while($looping)
Write-Host "Have a nice day :)"