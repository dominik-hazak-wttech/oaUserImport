$params = @{}
$params.firstName = Read-Host "Provide user's first name"
$params.lastName = Read-Host "Provide user's last name"
$params.userEmail = Read-Host "Provide user's email"

$parameters = @{}
$parameters.nickname = Read-Host "Provide data for username"
$parameters.line_managerid = Read-Host "Provide data for line manager (id)"
$parameters.departmentid = Read-Host "Provide data for department (id)"
$parameters.job_codeid = Read-Host "Provide data for job code (id)"
$parameters.UserCountry__c = Read-Host "Provide data for user country"
$parameters.EmploymentStatus__c = Read-Host "Provide data for employment status"
$parameters.Contract_type__c = Read-Host "Provide data for contract type"
$parameters.JobFunction__c = Read-Host "Provide data for functions for utilisation"
$parameters.Company__c = Read-Host "Provide data for company"
$parameters.UserLocation__c = Read-Host "Provide data for location"
$parameters.CoE__c = Read-Host "Provide data for CoE"
$parameters.Clan__c = ""
$parameters.Billability__c = Read-Host "Provide data for billability"
$parameters.VaultCode__c = ""
$parameters.active = "1"
$parameters.rate = ""
$parameters.password = Read-Host "Provide password for user"
$params.parameters = $parameters
$resp = $connector.SendRequest([OARequestType]::CreateUser,$params)
Write-Host ($resp.response.CreateUser.User | Format-List | Out-String)