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
$decision = Read-Host "You're about to read $($dataToProcess.Count) bookings. Are you sure? (type yes)"
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

    $bookingRead = @{}
    $bookingRead.type = "Booking"
    $bookingRead.method = "equal to"
    $bookingRead.customFields = $true
    $bookingRead.queryData = @{}
    $bookingRead.queryData.userid = $user.id
    $bookingRead.queryData.ownerid = $requester
    $bookingRead.queryData.projectid = $projectid
    $bookingRead.queryData.customerid = $customerid
    $bookingRead.queryData.startdate = @{day=$startDate.Day; month=$startDate.Month; year=$startDate.Year}
    $bookingRead.queryData.enddate = @{day=$endDate.Day; month=$endDate.Month; year=$endDate.Year}
    $bookingRead.queryData.as_percentage = $as_percentage
    if($as_percentage -eq 1){
        $bookingRead.queryData.percentage = $percentage
    }
    else{
        $bookingRead.queryData.hours = $numHours
    }
    $bookingRead.queryData.booking_typeid = $bookingTypeid
    $importList += $bookingRead
}

if($failState){
    Write-Error "Cannot progress with update due to errors above"
    break
}

$counter = [pscustomobject] @{ Value = 0 }
$groupSize = 1000
$groups = $importList | Group-Object -Property { [math]::Floor($counter.Value++ / $groupSize) }
if($groupSize -lt $importList.Count){
    Write-Host "List of bookings exceeds API limit for one request. Bookings are divided to $($groups.Count) groups"
}

if($skipped -gt 0){
    Write-Host "$skipped entries were skipped"
}

foreach($group in $groups){
    Write-Host "Processing group $($group.Name)"
    $params = @{}
    $params.readData = $group.Group
    if($validateOnly){
        Write-Host "Would send request to read group of $($group.Group.Count) bookings"
        Set-Content -Path "./request-$($group.Name).xml" -Value ($connector.SendRequest([OARequestType]::ReadBulk,$params,$true)).OuterXml
        continue
    }
    $resp = $connector.SendRequest([OARequestType]::ReadBulk,$params,$false)
    Write-Host "Saving transaction file"
    $transactionID = New-Guid
    Write-Host "Transaction ID: $transactionID"
    $successIDs = (($resp.response.Read | Where-Object {$_.status -eq "0"}).Booking | Select-Object -Property id).id
    Set-Content -Path "$logFolder/$transactionID.json" ($successIDs | ConvertTo-Json)
    $failedRequests = [System.Collections.ArrayList]@()
    if ($resp.response.Read.Count -eq 1){
        if($resp.response.Read.status -ne 0){
            $errorResp = $connector.SendRequest([OARequestType]::Read,@{limit="1";type="Error";method="equal to";queryData=@{code=($resp.response.Read[$i]).status}})
            $errorInfo = @{
                "ResourceID"=$params.readData.queryData.userid;
                "OwnerID"=$params.readData.queryData.ownerid;
                "ProjectID"=$params.readData.queryData.projectid;
                "Error code"=$resp.response.Read.status;
                "Error text"=$errorResp.response.Read.Error.comment;
                "OuterXml" = $resp.response.Read.OuterXml
            }
            $failedRequests.Add($errorInfo) | Out-Null
        }
    }
    else{
        for($i=0;$i -lt $resp.response.Read.Count; $i++){
            if(($resp.response.Read[$i]).status -ne 0){
                $errorResp = $connector.SendRequest([OARequestType]::Read,@{limit="1";type="Error";method="equal to";queryData=@{code=($resp.response.Read[$i]).status}})
                $errorInfo = @{
                    "ResourceID"=$params.readData[$i].queryData.userid;
                    "OwnerID"=$params.readData[$i].queryData.ownerid;
                    "ProjectID"=$params.readData[$i].queryData.projectid;
                    "Error code"=($resp.response.Read[$i]).status;
                    "Error text"=$errorResp.response.Read.Error.comment;
                    "OuterXml" = $resp.response.Read[$i].OuterXml
                }
                $failedRequests.Add($errorInfo) | Out-Null
            }
        }
    }
    Set-Content -Path "$logFolder/error-$transactionID.json" ($failedRequests | ConvertTo-Json | Out-String)
    Write-Host "Out of $($params.readData.Count):`n`t$($successIDs.Count) were read successfully`n`t$($failedRequests.Count) failed"
    Set-Content -Path "$logFolder/response-$transactionID.xml" $resp.response.OuterXml
}