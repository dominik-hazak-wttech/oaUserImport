$transactionID = Read-Host "Provide Transaction ID"
$idsToRemove = ((Get-Content -Path "$logFolder/$transactionID.txt") -split ';')
Write-Host "IDs to remove: $idsToRemove"
$resp = $connector.SendRequest([OARequestType]::DeleteUser,@{userIDs=$idsToRemove})
Set-Content -Path "$logFolder/deletedUsers-$transactionID.txt" "$($resp.response.Delete.Count) user accounts deleted `n`nResponse:`n$($resp.response.OuterXml)"