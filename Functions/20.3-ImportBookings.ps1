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
    if($project){
        Remove-Variable project
    }
    if($projectid){
        Remove-Variable projectid
    }
    if($customerid){
        Remove-Variable customerid
    }
    if($capability){
        Remove-Variable capability
    }
    if($careerStream){
        Remove-Variable careerStream
    }
    if($technicalStream){
        Remove-Variable technicalStream
    }
    if($careerLevel){
        Remove-Variable careerLevel
    }
    if($rateLevel){
        Remove-Variable rateLevel
    }

    if(-not $row."Resource" -or $row."Resource" -eq "0" -or $row."Resource".getType() -eq [OfficeOpenXml.ExcelErrorValue] -or $row."Resource" -eq "?"){
        Write-Warning "Booking is skipped as resource has value of $($row."Resource")"
        $skipped++
        continue
    }
    if(-not $row."Resource Manager" -or $row."Resource Manager" -eq "0" -or $row."Resource Manager".getType() -eq [OfficeOpenXml.ExcelErrorValue] -or $row."Resource Manager" -eq "?"){
        Write-Warning "Booking has empty resource manager (value of $($row."Resource Manager"))"
    }
    if(-not $row."Project" -or $row."Project" -eq "0" -or $row."Project".getType() -eq [OfficeOpenXml.ExcelErrorValue] -or $row."Project" -eq "?"){
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

    if(-not $row."Requester" -or $row."Requester" -eq "0" -or $row."Requester".getType() -eq [OfficeOpenXml.ExcelErrorValue] -or $row."Requester" -eq "?"){
        Write-Warning "Booking requester has value of $($row."Requester"). Booking will be reuested by scr_jml"
        $requester = 765 # scr_jml user ID
    }
    else{
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
        if($requesterFirst -and $requesterLast){
            $requester = ($checkedUsers | Where-Object { $_.addr.Address.first -eq $requesterFirst -and $_.addr.Address.last -eq $requesterLast }).id
        }
    }

    if($userFirst -and $userLast){
        $user = $checkedUsers | Where-Object { $_.addr.Address.first -eq $userFirst -and $_.addr.Address.last -eq $userLast }
    }
    elseif($userName){
        $user = $checkedUsers | Where-Object { $_.name -eq $userName }
    }

    $project = $checkedProjects | Where-Object { $_.name -eq $row."Project".Trim() }
    $projectid = $project.id
    $customerid = $project.customerid
    $startDate = $row."Start Date"
    $endDate = $row."End Date"
    $percentage = $row."Percentage of Time"
    $numHours = $row."Number of hours"
    $bookingTypeid = ($bookingTypes | Where-Object { $_.name -eq $row."Booking Type" }).id
    $resourceManager = $row."Resource Manager"
    $rate = $row."Sold Hourly Rate"
    $notes = $row."Notes"


    if (-not $user.id){
        if($importToSandbox){
            Write-Warning "User $($row."Resource") was not fund in OA. Booking will be skipped."
            $skipped++
            continue
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
            continue
        }
        else{
            Write-Error "Requester $($row."Requester") was not fund in OA."
            $failState = $true
        }
    }
    if(-not $projectid -or -not $customerid){
        if($importToSandbox){
            Write-Warning "Project $($row."Project") was not fund in OA. Booking will be skipped."
            $skipped++
            continue
        }
        else{
            Write-Error "Project $($row."Project") was not fund in OA."
            $failState = $true
        }
    }
    elseif($project.active -ne "1"){
        if($importToSandbox){
            Write-Warning "Project $($row."Project") is not active in OA. Booking will be skipped."
            $skipped++
            continue
        }
        else{
            Write-Error "Project $($row."Project") is not active in OA. Cannot continue."
            $failState = $true
        }
    }
    if($endDate -lt (Get-Date)){
        Write-Warning "Booking for user $($row."Resource") has end date in the past. Booking may fail to save."
    }

    if($percentage -gt 0){
        $as_percentage = "1"
        $percentage = $percentage * 100
    }
    elseif($numHours -gt 0){
        $as_percentage = "0"
    }
    else{
        Write-Error "Booking for user $($row."Resource") has neither percentage nor number of hours defined. Cannot continue."
        Write-Warning "percentage: $percentage, number of hours: $numHours"
        $failState = $true
        continue
    }

    $capability = $capabilities[$user."career_stream_user__c"]
    $careerStream = $user."career_stream_user__c"
    $technicalStream = $user."technical_stream_user__c"
    $careerLevel = $user."career_level_user__c"
    $rateLevel = $rateLevels[$user."UserCountry__c"]

    $bookingObj = @{}
    $bookingObj.type = "Booking"
    $bookingObj.dataToAdd = @{}
    $bookingObj.dataToAdd.userid = $user.id
    $bookingObj.dataToAdd.ownerid = $requester
    $bookingObj.dataToAdd.projectid = $projectid
    $bookingObj.dataToAdd.customerid = $customerid
    $bookingObj.dataToAdd.startdate = @{day=$startDate.Day; month=$startDate.Month; year=$startDate.Year}
    $bookingObj.dataToAdd.enddate = @{day=$endDate.Day; month=$endDate.Month; year=$endDate.Year}
    $bookingObj.dataToAdd.as_percentage = $as_percentage
    if($as_percentage -eq 1){
        $bookingObj.dataToAdd.percentage = $percentage
    }
    else{
        $bookingObj.dataToAdd.hours = $numHours
    }
    $bookingObj.dataToAdd.booking_typeid = $bookingTypeid
    $bookingObj.dataToAdd.resource_manager__c = $resourceManager
    $bookingObj.dataToAdd.sold_rate__c = $rate
    $bookingObj.dataToAdd.notes = $notes
    $bookingObj.dataToAdd.division__c = $capability
    $bookingObj.dataToAdd.career_slice__c = $careerStream
    $bookingObj.dataToAdd.technical_stream__c = $technicalStream
    $bookingObj.dataToAdd.career_level__c = $careerLevel
    $bookingObj.dataToAdd.career_level_hidden__c = $careerLevel
    $bookingObj.dataToAdd.rate_level__c = $rateLevel
    $bookingObj.dataToAdd.rate_level_hidden__c = $rateLevel
    $importList += $bookingObj
}

if($failState){
    Write-Error "Cannot progress with update due to errors above"
    break
}

$counter = [pscustomobject] @{ Value = 0 }
$groupSize = 1000
$groups = $importList | Group-Object -Property { [math]::Floor($counter.Value++ / $groupSize) }
if($groupSize -lt $importList.Count){
    Write-Host "List of users exceeds API limit for one request. Users are divided to $($groups.Count) groups"
}

if($skipped -gt 0){
    Write-Host "$skipped entries were skipped"
}

foreach($group in $groups){
    Write-Host "Processing group $($group.Name)"
    $params = @{}
    $params.addRequests = $group.Group
    if($validateOnly){
        Write-Host "Would send request to create group of $($group.Group.Count) bookings"
        Set-Content -Path "./request-$($group.Name).xml" -Value ($connector.SendRequest([OARequestType]::AddBulk,$params,$true)).OuterXml
        continue
    }
    $resp = $connector.SendRequest([OARequestType]::AddBulk,$params,$false)
    Write-Host "Saving transaction file"
    $transactionID = New-Guid
    Write-Host "Transaction ID: $transactionID"
    $successIDs = (($resp.response.Add | Where-Object {$_.status -eq "0"}).Booking | Select-Object -Property id).id
    Set-Content -Path "$logFolder/$transactionID.json" ($successIDs | ConvertTo-Json)
    $failedRequests = [System.Collections.ArrayList]@()
    if ($resp.response.Add.Count -eq 1){
        if($resp.response.Add.status -ne 0){
            $errorResp = $connector.SendRequest([OARequestType]::Read,@{limit="1";type="Error";method="equal to";queryData=@{code=($resp.response.Add[$i]).status}})
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
            if(($resp.response.Add[$i]).status -ne 0){
                $errorResp = $connector.SendRequest([OARequestType]::Read,@{limit="1";type="Error";method="equal to";queryData=@{code=($resp.response.Add[$i]).status}})
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
    Set-Content -Path "$logFolder/error-$transactionID.json" ($failedRequests | ConvertTo-Json | Out-String)
    Write-Host "Out of $($params.addRequests.Count):`n`t$($successIDs.Count) were created successfully`n`t$($failedRequests.Count) failed"
    Set-Content -Path "$logFolder/response-$transactionID.xml" $resp.response.OuterXml
}