$usersToDisable = (Get-Content -Path "$logFolder/UsersToDisable.txt") -split ";"
$modRequests = @()
foreach($id in $usersToDisable){
    $modRequest = @{
        type = "User";
        id = $id;
        dataToUpdate = @{
            active = "1"
        }
    }
    $modRequests += $modRequest
}
$resp = $connector.SendRequest([OARequestType]::ModifyBulk,@{modifyRequests = $modRequests})
Write-Host $resp.response.OuterXml