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