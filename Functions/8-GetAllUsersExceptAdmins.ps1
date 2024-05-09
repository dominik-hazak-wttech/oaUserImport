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