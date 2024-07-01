$usersToUpdate = @(
    @{"User ID" = "alex.nikitits@vml.com"},
    @{"User ID" = "bob.hai@vml.com"},
    @{"User ID" = "chitra.pattarveedu@vml.com"},
    @{"User ID" = "diego.neves@vml.com"},
    @{"User ID" = "jim.sun@vml.com"},
    @{"User ID" = "mark.horton@vml.com"},
    @{"User ID" = "murali.chava@vml.com"},
    @{"User ID" = "peter.gao@vml.com"},
    @{"User ID" = "sam.el-yafi@vml.com"},
    @{"User ID" = "scott.harperashton@vml.com"},
    @{"User ID" = "timothy.shipman@vml.com"},
    @{"User ID" = "victor.yan@vml.com"},
    @{"User ID" = "alessandro.fulgenzi@vml.com"},
    @{"User ID" = "joao.soaresbranco@vml.com"},
    @{"User ID" = "navin.shet@vml.com"},
    @{"User ID" = "ryan.hutchings@vml.com"},
    @{"User ID" = "tarik.canessawright@vml.com"},
    @{"User ID" = "venu.vasudevan@vml.com"},
    @{"User ID" = "samuel.suder@vml.com"},
    @{"User ID" = "carolina.santos@vml.com"},
    @{"User ID" = "kieren.hinch@vml.com"},
    @{"User ID" = "robert.good@vml.com"},
    @{"User ID" = "beth.marchant@vml.com"},
    @{"User ID" = "darren.cooper@vml.com"},
    @{"User ID" = "jordan.cox@vml.com"},
    @{"User ID" = "richard.desouzafigueiredo@vml.com"},
    @{"User ID" = "sean.varnham@vml.com"},
    @{"User ID" = "timothy.gibson@vml.com"},
    @{"User ID" = "prabhu.malayan@vml.com"},
    @{"User ID" = "ashwini.chikane@vml.com"},
    @{"User ID" = "jayesh.karadia@vml.com"},
    @{"User ID" = "vishal.singh@vml.com"},
    @{"User ID" = "gaurav.pabreja@vml.com"},
    @{"User ID" = "mohammed.alishaik@vml.com"},
    @{"User ID" = "mary.thumma@vml.com"},
    @{"User ID" = "puneet.garg@vml.com"},
    @{"User ID" = "sewa.kc@vml.com"},
    @{"User ID" = "tom.hyland@vml.com"},
    @{"User ID" = "mel.mcdonald@vml.com"},
    @{"User ID" = "kevin.williams@vml.com"},
    @{"User ID" = "kuldip.buttar@vml.com"},
    @{"User ID" = "sankar.krishnamoorthi@vml.com"},
    @{"User ID" = "scott.fuller@vml.com"},
    @{"User ID" = "maja.zor@vml.com"},
    @{"User ID" = "paulina.ratomska@vml.com"},
    @{"User ID" = "abhishek01.kumar@vml.com"},
    @{"User ID" = "durga.gavara@vml.com"},
    @{"User ID" = "emmanuel.ekewenu@vml.com"},
    @{"User ID" = "philippa.hodge@vml.com"},
    @{"User ID" = "siim.valner@vml.com"},
    @{"User ID" = "venkata.krishnagomatam@vml.com"},
    @{"User ID" = "max.hamilton@vml.com"},
    @{"User ID" = "avneet.mudhar@vml.com"},
    @{"User ID" = "erik.benchak@vml.com"},
    @{"User ID" = "kevin.french@vml.com"},
    @{"User ID" = "bhargavi.pasapula@vml.com"},
    @{"User ID" = "charles.henry@vml.com"},
    @{"User ID" = "dan.lockett@vml.com"},
    @{"User ID" = "marcin.jedynak@vml.com"},
    @{"User ID" = "maria.ilina@vml.com"},
    @{"User ID" = "miro.kmet@vml.com"},
    @{"User ID" = "nithya.vittal@vml.com"},
    @{"User ID" = "pannerselvam.pazhani@vml.com"},
    @{"User ID" = "prachi.sahansarval@vml.com"},
    @{"User ID" = "rani.jose@vml.com"},
    @{"User ID" = "suresh.gangodawila@vml.com"},
    @{"User ID" = "sushil.tiwari@vml.com"},
    @{"User ID" = "benny.song@vml.com"}
)
if(-not $usersToUpdate){
    Write-Host "You need to load data first" -ForegroundColor Red
    break
}

$OAStatusReady = "READY FOR IMPORT"
$OACleaningStatusValidate = "NEEDS VALIDATION"
$OAStatusImported = "IMPORTED"
$OAStatusSkip = "SKIP"
$OARegularAccount = "Regular"
$OASkeletonAccount = "Skeleton"

$dataToProcess = $usersToUpdate
# $dataToProcess = $dataToProcess | Where-Object {$_."OA user TYPE" -eq $OARegularAccount}
$decision = Read-Host "You're about to udpate $($dataToProcess.Count) accounts. Are you sure? (type yes)"
if($decision.ToLower() -ne "yes"){
    Write-Host "User account creation aborted"
    break
}
$importList = @()
foreach($row in $dataToProcess){
    if ($row."User ID" -eq "" -or $null -eq $row."User ID"){
        Write-Error "$($row.Email) is missing User ID value!"
        $failState = $true
    }
    if (-not $userDataDictionary[$row."User ID"]){
        Write-Warning "$($row."User ID") is missing id in dictionary! User will be skipped!"
        continue
    }
    $userObj = @{}
    $userObj.id = $userDataDictionary[$row."User ID"]
    $userObj.type = "User"
    $userObj.dataToUpdate = @{}
    $userObj.dataToUpdate.saml_auth__c = "1"
    $importList += $userObj
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
foreach($group in $groups){
    $params = @{}
    $params.modifyRequests = $group.Group
    if($validateOnly){
        Write-Host "Would send request to edit group of $($group.Group.Count) accounts"
        Set-Content -Path ./request.xml -Value ($connector.SendRequest([OARequestType]::ModifyBulk,$params,$true)).OuterXml
        continue
    }
    $resp = $connector.SendRequest([OARequestType]::ModifyBulk,$params,$false)
    Write-Host "Saving transaction file"
    $transactionID = New-Guid
    $successIDs = (($resp.response.Modify | Where-Object {$_.status -eq "0"}).User | Select-Object -Property id).id
    $modifiedUserNames = ($resp.response.Modify.User | Select-Object -Property nickname).nickname
    Set-Content -Path "$logFolder/$transactionID.txt" ($successIDs -join ';')
    Set-Content -Path "$logFolder/modifiedUsers-$transactionID.txt" ($modifiedUserNames -join ';')
    $failedRequests = [System.Collections.ArrayList]@()
    if ($resp.response.Modify.Count -eq 1){
        if($resp.response.Modify.status -ne 0){
            $errorResp = $connector.SendRequest([OARequestType]::Read,@{limit="1";type="Error";method="equal to";queryData=@{code=($resp.response.Modify[$i]).status}})
            $errorInfo = @{
                "ID" = $params.modifyRequests.id;
                "Error code"=$resp.response.Modify.status;
                "Error text"=$errorResp.response.Read.Error.text;
                "OuterXml" = $resp.response.Modify.OuterXml
            }
            $failedRequests.Add($errorInfo)
        }
    }
    else{
        for($i=0;$i -lt $resp.response.Modify.Count; $i++){
            if(($resp.response.Modify[$i]).status -ne 0){
                $errorResp = $connector.SendRequest([OARequestType]::Read,@{limit="1";type="Error";method="equal to";queryData=@{code=($resp.response.Modify[$i]).status}})
                $errorInfo = @{
                    "ID" = $params.modifyRequests[$i].id;
                    "Error code"=($resp.response.Modify[$i]).status;
                    "Error text"=$errorResp.response.Read.Error.text;
                    "OuterXml" = ($resp.response.Modify[$1]).OuterXml
                }
                $failedRequests.Add($errorInfo)
            }
        }
    }
    Set-Content -Path "$logFolder/error-$transactionID.txt" ($failedRequests | Format-List | Out-String)
    Write-Host "Out of $($params.usersData.Count):`n`t$($successIDs.Count) were modified successfully`n`t$($failedRequests.Count) failed"
    Write-Host "Transaction ID: $transactionID"
    Set-Content -Path "$logFolder/response-$transactionID.xml" $resp.OuterXml
}