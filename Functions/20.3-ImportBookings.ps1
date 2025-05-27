if(-not $bulkData){
    Write-Host "You need to load data first" -ForegroundColor Red
    break
}

if(-not $checkedUsers){
    Write-Host "You need to read users first" -ForegroundColor Red
    break
}

if(-not $checkedProjects){
    Write-Host "You need to read projects first" -ForegroundColor Red
    break
}

if(-not $bookingTypes){
    Write-Host "You need to read booking types first" -ForegroundColor Red
    break
}

$dataToProcess = $bulkData
$decision = Read-Host "You're about to add $($dataToProcess.Count) bookings. Are you sure? (type yes)"
if($decision.ToLower() -ne "yes"){
    Write-Host "Operation aborted"
    break
}

$skipped=0
$importList = @()
$failState = $false
foreach($row in $dataToProcess){
    if($name){
        Remove-Variable name
    }
    if($first){
        Remove-Variable first
    }
    if($last){
        Remove-Variable last
    }
    if($userFirst){
        Remove-Variable userFirst
    }
    if($userLast){
        Remove-Variable userLast
    }
    if($userName){
        Remove-Variable userName
    }
    if($requesterFirst){
        Remove-Variable requesterFirst
    }
    if($requesterLast){
        Remove-Variable requesterLast
    }
    if($user){
        Remove-Variable user
    }
    if($requester){
        Remove-Variable requester
    }
    if($projectid){
        Remove-Variable projectid
    }
    if($customerid){
        Remove-Variable customerid
    }

    if($row."Resource" -eq "0" -or $row."Resource".getType() -eq [OfficeOpenXml.ExcelErrorValue] -or -not $row."Resource"){
        Write-Warning "Booking is skipped as resource has value of $($row."Resource")"
        $skipped++
        continue
    }
    if($row."Requester" -eq "0" -or $row."Requester".getType() -eq [OfficeOpenXml.ExcelErrorValue] -or -not $row."Requester"){
        Write-Warning "Booking is skipped as requester has value of $($row."Requester")"
        $skipped++
        continue
    }
    if($row."Resource Manager" -eq "0" -or $row."Resource Manager".getType() -eq [OfficeOpenXml.ExcelErrorValue] -or -not $row."Resource Manager"){
        Write-Warning "Booking is skipped as resource manager has value of $($row."Resource Manager")"
        $skipped++
        continue
    }
    if($row."Project" -eq "0" -or $row."Project".getType() -eq [OfficeOpenXml.ExcelErrorValue] -or -not $row."Project"){
        Write-Warning "Booking is skipped as project has value of $($row."Project")"
        $skipped++
        continue
    }

    $userSplit = $row."Resource" -split ","
    if ($userSplit.Length -eq 2){
        $userFirst = $userSplit[1].Trim()
        $userLast = $userSplit[0].Trim()
    }
    else{
        $userName = $row."Resource".Trim()
    }

    $requesterSplit = $row."Requester" -split ","
    if ($requesterSplit.Length -eq 2){
        $requesterFirst = $requesterSplit[1].Trim()
        $requesterLast = $requesterSplit[0].Trim()
    }
    else{
        Write-Error "Requester value doesn't split to two parts: $($row."Requester")"
        $failState = $true
        continue
    }

    if($userFirst -and $userLast){
        $user = ($checkedUsers | Where-Object { $_.addr.Address.first -eq $userFirst -and $_.addr.Address.last -eq $userLast }).id
    }
    elseif($userName){
        $user = ($checkedUsers | Where-Object { $_.name -eq $userName }).id
    }

    if($requesterFirst -and $requesterLast){
        $requester = ($checkedUsers | Where-Object { $_.addr.Address.first -eq $requesterFirst -and $_.addr.Address.last -eq $requesterLast }).id
    }

    $projectid = ($checkedProjects | Where-Object { $_.name -eq $row."Project".Trim() }).id
    $customerid = ($checkedProjects | Where-Object { $_.name -eq $row."Project".Trim() }).customerid
    $startDate = $row."Start Date"
    $endDate = $row."End Date"
    $as_percentage = "1"
    $percentage = $row."Percentage of Time"
    $bookingTypeid = ($bookingTypes | Where-Object { $_.name -eq $row."Booking Type" }).id
    $resourceManager = $row."Resource Manager"
    $rate = $row."Sold Hourly Rate"
    $notes = $row."Notes"


    if (-not $user){
        if($importToSandbox){
            Write-Warning "User $($row."Resource") was not fund in OA. Booking will be skipped."
            $skipped++
        }
        else{
            Write-Error "User $($row."Resource") was not fund in OA."
            $failState = $true
        }
        
    }
    if(-not $requester){
        if($importToSandbox){
            Write-Warning "Requester $($row."Requester") was not fund in OA. Booking will be skipped."
            $skipped++
        }
        else{
            Write-Error "Requester $($row."Requester") was not fund in OA."
            $failState = $true
        }
    }
    if(-not $projectid -or -not $customerid){
        Write-Error "Project $($row."Project") was not fund in OA."
        $failState = $true
    }
    $bookingObj = @{}
    $bookingObj.type = "Booking"
    $bookingObj.dataToAdd = @{}
    $bookingObj.dataToAdd.userid = $user
    $bookingObj.dataToAdd.ownerid = $requester
    $bookingObj.dataToAdd.projectid = $projectid
    $bookingObj.dataToAdd.customerid = $customerid
    $bookingObj.dataToAdd.startdate = @{day=$startDate.Day; month=$startDate.Month; year=$startDate.Year}
    $bookingObj.dataToAdd.enddate = @{day=$endDate.Day; month=$endDate.Month; year=$endDate.Year}
    $bookingObj.dataToAdd.as_percentage = $as_percentage
    $bookingObj.dataToAdd.percentage = $percentage
    $bookingObj.dataToAdd.booking_typeid = $bookingTypeid
    $bookingObj.dataToAdd.resource_manager__c = $resourceManager
    $bookingObj.dataToAdd.sold_rate__c = $rate
    $bookingObj.dataToAdd.notes = $notes
    $importList += $bookingObj
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

if($skipped -gt 0){
    Write-Host "$skipped entries were skipped"
}

foreach($group in $groups){
    $params = @{}
    $params.addRequests = $group.Group
    if($validateOnly){
        Write-Host "Would send request to create group of $($group.Group.Count) bookings"
        Set-Content -Path ./request.xml -Value ($connector.SendRequest([OARequestType]::AddBulk,$params,$true)).OuterXml
        continue
    }
    $resp = $connector.SendRequest([OARequestType]::AddBulk,$params,$false)
    Write-Host "Saving transaction file"
    $transactionID = New-Guid
    Write-Host "Transaction ID: $transactionID"
    $successIDs = (($resp.response.Add | Where-Object {$_.status -eq "0"}).Booking | Select-Object -Property id).id
    Set-Content -Path "$logFolder/$transactionID.txt" ($successIDs -join ';')
    $failedRequests = [System.Collections.ArrayList]@()
    if ($resp.response.Add.Count -eq 1){
        if($resp.response.Add.status -ne 0){
            $errorResp = $connector.SendRequest([OARequestType]::Read,@{limit="1";type="Error";method="equal to";queryData=@{code=($resp.response.Modify[$i]).status}})
            $errorInfo = @{
                "ResourceID"=$params.addRequests.dataToAdd.userid;
                "OwnerID"=$params.addRequests.dataToAdd.ownerid;
                "ProjectID"=$params.addRequests.dataToAdd.projectid;
                "Error code"=$resp.response.Add.status;
                "Error text"=$errorResp.response.Read.Error.comment;
                "OuterXml" = $resp.response.Add.OuterXml
            }
            $failedRequests.Add($errorInfo) | Out-Null
        }
    }
    else{
        for($i=0;$i -lt $resp.response.Add.Count; $i++){
            if(($resp.response.Modify[$i]).status -ne 0){
                $errorResp = $connector.SendRequest([OARequestType]::Read,@{limit="1";type="Error";method="equal to";queryData=@{code=($resp.response.Modify[$i]).status}})
                $errorInfo = @{
                    "ResourceID"=$params.addRequests[$i].dataToAdd.userid;
                    "OwnerID"=$params.addRequests[$i].dataToAdd.ownerid;
                    "ProjectID"=$params.addRequests[$i].dataToAdd.projectid;
                    "Error code"=($resp.response.Add[$i]).status;
                    "Error text"=$errorResp.response.Read.Error.comment;
                    "OuterXml" = $resp.response.Add[$i].OuterXml
                }
                $failedRequests.Add($errorInfo) | Out-Null
            }
        }
    }
    Set-Content -Path "$logFolder/error-$transactionID.txt" ($failedRequests | Format-List | Out-String)
    Write-Host "Out of $($params.addRequests.Count):`n`t$($successIDs.Count) were created successfully`n`t$($failedRequests.Count) failed"
    Set-Content -Path "$logFolder/response-$transactionID.xml" $resp.response.OuterXml
}