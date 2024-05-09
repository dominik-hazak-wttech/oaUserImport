if(-not $bulkData){
    Write-Host "You need to load data first" -ForegroundColor Red
    break
}
$dataToProcess = $bulkData
$dataToProcess = $dataToProcess | Where-Object {$_."OA Import Status" -eq "READY FOR IMPORT"}
$decision = Read-Host "You're about to create $($dataToProcess.Count) accounts. Are you sure? (type yes)"
if($decision.ToLower() -ne "yes"){
    Write-Host "User account creation aborted"
    break
}
$importList = @()
foreach($row in $dataToProcess){
    $userObj = @{}
    $userObj.firstName = $row."First Name"
    $userObj.lastName = $row."Last Name"
    $userObj.userEmail = $row.Email
    $userObj.parameters = @{}
    $userObj.parameters.nickname = $row."User ID"
    $userObj.parameters.line_managerid = $row.Manager
    $userObj.parameters.departmentid = $row.Department
    $userObj.parameters.job_codeid = $row."Job code"
    $userObj.parameters.UserCountry__c = $row."User Country"
    $userObj.parameters.EmploymentStatus__c = $row."Employment status"
    $userObj.parameters.Contract_type__c = $row."Contract type"
    $userObj.parameters.JobFunction__c = $row."Functions For Utilisation"
    $userObj.parameters.Company__c = $row.Company
    $userObj.parameters.UserLocation__c = $row.Location
    $userObj.parameters.CoE__c = $row.CoE
    $userObj.parameters.Clan__c = $row.Clan
    $userObj.parameters.Billability__c = $row.Billability
    $userObj.parameters.VaultCode__c = $row.VaultCode
    $userObj.parameters.active = ($row."Is Active" -eq "Active") ? 1 : 0
    $userObj.parameters.rate = $row.Cost
    $userObj.parameters.password = $row.Password
    $importList += $userObj
}
$counter = [pscustomobject] @{ Value = 0 }
$groupSize = 999
$groups = $importList | Group-Object -Property { [math]::Floor($counter.Value++ / $groupSize) }
if($groupSize -lt $importList.Count){
    Write-Host "List of users exceeds API limit for one request. Users are divided to $($groups.Count) groups"
}
foreach($group in $groups){
    $request = $connector.GenerateCreateUserBulkRequest($group.Group)
    Write-Host ($request.OuterXml)
}