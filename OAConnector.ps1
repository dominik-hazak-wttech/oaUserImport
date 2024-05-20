enum OARequestType{
    Read
    CreateUser
    CreateUserBulk
    Whoami
    Time
    Auth
    DeleteUser
    Modify
    ModifyBulk
    AddBulk
}

class OAConnector{
    
    [hashtable] $authTypes = @{
        "0" = "Password";
        "1" = "Token"
    }

    [string] $namespace
    [string] $apiKey
    [hashtable] $OACredentials
    [xml] $xmlDocument
    [string] $OAApiEndpoint = "https://cognifide-ltd.app.sandbox.openair.com/api.pl"
    [array] $userCreationRequired

    OAConnector([string]$namespace, [string]$apiKey, [string]$apiUrl){
        $this.namespace = $namespace
        $this.apiKey = $apiKey
        if($null -ne $apiUrl){
            $this.OAApiEndpoint = $apiUrl
        }

        $authDec = ""
        do{
            $authDec = Read-Host "What type of authentication you want to use? (0 - Password, 1 - Token)"
        }
        while ($authDec -notin $this.authTypes.Keys)
        if($this.authTypes[$authDec] -eq "Password"){
            Write-Host "Provide OpenAir credentials"
            $company = Read-Host -Prompt "Company"
            $login = Read-Host -Prompt "Login"
            $password = Read-Host -Prompt "Password" -MaskInput
            $this.OACredentials = @{"Company" = $company; "User" = $login; "Password" = $password}
        }
        elseif ($this.authTypes[$authDec] -eq "Token"){
            $accessToken = Read-Host -Prompt "Provide OpenAir access token"
            $this.OACredentials = @{"access_token" = $accessToken}
        }
        $this.xmlDocument = $this.GenerateRequestDocument()
    }
    
    OAConnector([string]$namespace, [string]$apiKey, [string]$apiUrl, [string]$accessToken){
        $this.namespace = $namespace
        $this.apiKey = $apiKey
        if($null -ne $apiUrl){
            $this.OAApiEndpoint = $apiUrl
        }

        Write-Host "Provide OpenAir credentials"
        $newToken = Read-Host -Prompt "Access Token (currently: $($accessToken[0..8])...)"
        if($newToken){
            $accessToken = $newToken
        }
        $this.OACredentials = @{"access_token" = $accessToken}
        $this.xmlDocument = $this.GenerateRequestDocument()
    }

    OAConnector([string]$namespace, [string]$apiKey, [string]$apiUrl, [string]$company, [string]$login){
        $this.namespace = $namespace
        $this.apiKey = $apiKey
        if($null -ne $apiUrl){
            $this.OAApiEndpoint = $apiUrl
        }

        Write-Host "Provide OpenAir credentials"
        $newCompany = Read-Host -Prompt "Company (currently: $company)"
        if($newCompany){
            $company = $newCompany
        }
        $newLogin = Read-Host -Prompt "Login (currently: $login)"
        if($newLogin)
        {
            $login = $newLogin
        }
        $password = Read-Host -Prompt "Password" -MaskInput
        $this.OACredentials = @{"Company" = $company; "User" = $login; "Password" = $password}
        $this.xmlDocument = $this.GenerateRequestDocument()
    }
    
    [xml] GenerateRequestDocument(){
        [xml] $requestDocument = New-Object -TypeName xml
        $xmlDeclaration = $requestDocument.CreateXmlDeclaration("1.0","UTF-8","yes")
        
        $request = $requestDocument.CreateElement("request")
        $request.SetAttribute("API_ver","1.0")
        $request.SetAttribute("client","OAImporter")
        $request.SetAttribute("client_ver","1.0")
        $request.SetAttribute("namespace","$($this.namespace)")
        $request.SetAttribute("key","$($this.apiKey)")

        $requestDocument.AppendChild($request)
        $requestDocument.InsertBefore($xmlDeclaration,$request)
        return $requestDocument
    }

    [System.Xml.XmlElement] GenerateAuthElement([xml] $xml){
        $authLogin = $xml.CreateElement("Login")
        if($this.OACredentials.access_token){
            $authToken = $xml.CreateElement("access_token")
            $authToken.InnerText = $this.OACredentials.access_token
            $authLogin.AppendChild($authToken)
        }
        else{
            $authCompany = $xml.CreateElement("company")
            $authCompany.InnerText = $this.OACredentials.Company
            
            $authUser = $xml.CreateElement("user")
            $authUser.InnerText = $this.OACredentials.User
            
            $authPass = $xml.CreateElement("password")
            $authPass.InnerText = $this.OACredentials.Password
            
            $authLogin.AppendChild($authCompany)
            $authLogin.AppendChild($authUser)
            $authLogin.AppendChild($authPass)
        }
        
        $auth = $xml.CreateElement("Auth")
        $auth.AppendChild($authLogin)

        return $auth
    }
    
    [System.Xml.XmlElement] GenerateReadElement([xml] $xml, [string]$type, [string]$method, [hashtable]$queryData, [boolean]$customFields, [int]$limit){
        $typeElement = $xml.CreateElement($type)
        foreach ($key in $queryData.Keys){
            $queryElement = $xml.CreateElement($key)
            $queryElement.InnerText = $queryData.$key
            $typeElement.AppendChild($queryElement)
        }
        Write-Host "Emable custom current value: $customFields"
        $readElement = $xml.CreateElement("Read")
        $readElement.SetAttribute("type","$type")
        $readElement.SetAttribute("method","$method")
        $readElement.SetAttribute("limit","$limit")
        $readElement.SetAttribute("enable_custom",$customFields ? "1" : "0")
        $readElement.AppendChild($typeElement)
        return $readElement
    }

    [System.Xml.XmlElement] GenerateTimeElement([xml]$xml){
        $timeElement = $xml.CreateElement("Time")
        $timeElement.InnerText = " "
        return $timeElement
    }
    [System.Xml.XmlElement] GenerateWhoamiElement([xml]$xml){
        $timeElement = $xml.CreateElement("Whoami")
        $timeElement.InnerText = " "
        return $timeElement
    }

    [System.Xml.XmlElement] GenerateCreateUserElement([xml]$xml, [string]$firstName, [string]$lastName, [string]$userEmail, [hashtable]$parameters){
        if(-not $parameters.Keys -contains "nickname"){ Write-Host -ForegroundColor Red "Parameters are missing nickname!";return $null }
        if(-not $parameters.Keys -contains "line_managerid"){ Write-Host -ForegroundColor Red "Parameters are missing line_managerid!";return $null }
        if(-not $parameters.Keys -contains "departmentid"){ Write-Host -ForegroundColor Red "Parameters are missing departmentid!";return $null }
        if(-not $parameters.Keys -contains "job_codeid"){ Write-Host -ForegroundColor Red "Parameters are missing job_codeid!";return $null }
        if(-not $parameters.Keys -contains "UserCountry__c"){ Write-Host -ForegroundColor Red "Parameters are missing UserCountry__c!";return $null }
        if(-not $parameters.Keys -contains "EmploymentStatus__c"){ Write-Host -ForegroundColor Red "Parameters are missing EmploymentStatus__c!";return $null }
        if(-not $parameters.Keys -contains "Contract_type__c"){ Write-Host -ForegroundColor Red "Parameters are missing Contract_type__c!";return $null }
        if(-not $parameters.Keys -contains "JobFunction__c"){ Write-Host -ForegroundColor Red "Parameters are missing JobFunction__c!";return $null }
        if(-not $parameters.Keys -contains "Company__c"){ Write-Host -ForegroundColor Red "Parameters are missing Company__c!";return $null }
        if(-not $parameters.Keys -contains "UserLocation__c"){ Write-Host -ForegroundColor Red "Parameters are missing UserLocation__c!";return $null }
        if(-not $parameters.Keys -contains "CoE__c"){ Write-Host -ForegroundColor Red "Parameters are missing CoE__c!";return $null }
        if(-not $parameters.Keys -contains "Clan__c"){ Write-Host -ForegroundColor Red "Parameters are missing Clan__c!";return $null }
        if(-not $parameters.Keys -contains "Billability__c"){ Write-Host -ForegroundColor Red "Parameters are missing Billability__c!";return $null }
        if(-not $parameters.Keys -contains "VaultCode__c"){ Write-Host -ForegroundColor Red "Parameters are missing VaultCode__c!";return $null }
        if(-not $parameters.Keys -contains "active"){ Write-Host -ForegroundColor Red "Parameters are missing active!";return $null }
        if(-not $parameters.Keys -contains "rate"){ Write-Host -ForegroundColor Red "Parameters are missing rate!";return $null }

        $createUserElement = $xml.CreateElement("CreateUser")
        $createUserElement.SetAttribute("enable_custom","1")
        
        $nicknameElement = $xml.CreateElement("nickname")
        $nicknameElement.InnerText = "$($this.OACredentials.Company)"
        $companyElement = $xml.CreateElement("Company") 
        $companyElement.AppendChild($nicknameElement)
        $createUserElement.AppendChild($companyElement)

        $userElement = $xml.CreateElement("User")
        $addressElement = $xml.CreateElement("Address")
        $emailElement = $xml.CreateElement("email")
        $emailElement.InnerText = $userEmail
        $addressElement.AppendChild($emailElement)
        
        $firstNameElement = $xml.CreateElement("first")
        $firstNameElement.InnerText = $firstName
        $addressElement.AppendChild($firstNameElement)

        $lastNameElement = $xml.CreateElement("last")
        $lastNameElement.InnerText = $lastName
        $addressElement.AppendChild($lastNameElement)

        $addrElement = $xml.CreateElement("addr")
        $addrElement.AppendChild($addressElement)
        $userElement.AppendChild($addrElement)

        foreach ($key in $parameters.Keys){
            $parameterElement = $xml.CreateElement($key)
            $parameterElement.InnerText = $parameters[$key]
            $userElement.AppendChild($parameterElement)
        }
        $flagElement = $xml.CreateElement("Flag")
        $flagNameElement = $xml.CreateElement("name")
        $flagNameElement.InnerText = "ta_timesheet_required"
        $flagElement.AppendChild($flagNameElement)

        $flagSettingElement = $xml.CreateElement("setting")
        $flagSettingElement.InnerText = "0"
        $flagElement.AppendChild($flagSettingElement)

        $flagsElement = $xml.CreateELement("flags")
        $flagsElement.AppendChild($flagElement)
        $userElement.AppendChild($flagsElement)
        
        $createUserElement.AppendChild($userElement)
        return $createUserElement
    }

    [System.Xml.XmlElement] GenerateDeleteUserElement([xml]$xml, [string]$id){
        $idElement = $xml.CreateElement("id")
        $idElement.InnerText = $id

        $userElement = $xml.CreateElement("User")
        $userElement.AppendChild($idElement)

        $deleteElement = $xml.CreateElement("Delete")
        $deleteElement.SetAttribute("type","User")
        $deleteElement.AppendChild($userElement)
        return $deleteElement
    }
    
    [System.Xml.XmlElement] GenerateModifyElement([xml]$xml, [string]$type, [string]$id, [hashtable]$dataToUpdate){
        $idElement = $xml.CreateElement("id")
        $idElement.InnerText = $id

        $typeElement = $xml.CreateElement($type)
        $typeElement.AppendChild($idElement)

        foreach($key in $dataToUpdate.Keys){
            $attrElement = $xml.CreateElement($key)
            $attrElement.InnerText = $dataToUpdate.$key
            $typeElement.AppendChild($attrElement)
        }

        $modifyElement = $xml.CreateElement("Modify")
        $modifyElement.SetAttribute("type",$type)
        $modifyElement.SetAttribute("enable_custom","1")
        $modifyElement.AppendChild($typeElement)
        return $modifyElement
    }

    [System.Xml.XmlElement] GenerateAddElement([xml]$xml, [string]$type, [hashtable]$dataToAdd){
        $typeElement = $xml.CreateElement($type)

        foreach($key in $dataToAdd.Keys){
            $attrElement = $xml.CreateElement($key)
            if($dataToAdd.$key){
                $attrElement.InnerText = $dataToAdd.$key
            }
            $typeElement.AppendChild($attrElement)
        }

        $modifyElement = $xml.CreateElement("Add")
        $modifyElement.SetAttribute("type",$type)
        $modifyElement.SetAttribute("enable_custom","1")
        $modifyElement.AppendChild($typeElement)
        return $modifyElement
    }

    [xml] GenerateReadRequest([string]$type, [string]$method, [hashtable]$queryData, [boolean]$customFields, [int]$limit){
        $request = $this.xmlDocument.Clone()
        $authElement = $this.GenerateAuthElement($request)
        $request.DocumentElement.AppendChild($authElement)
        $readElement = $this.GenerateReadElement($request, $type, $method, $queryData, $customFields, $limit)
        $request.DocumentElement.AppendChild($readElement)
        return $request
    }
    
    [xml] GenerateReadRequest([string]$type, [string]$method, [hashtable]$queryData){
        return $this.GenerateReadRequest($type, $method, $queryData, $false, 10)
    }

    [xml] GenerateReadRequest([string]$type, [string]$method, [hashtable]$queryData, [int]$limit){
        return $this.GenerateReadRequest($type, $method, $queryData, $false, $limit)
    }

    [xml] GenerateReadRequest([string]$type, [string]$method, [hashtable]$queryData, [boolean]$customFields){
        return $this.GenerateReadRequest($type, $method, $queryData, $customFields, 10)
    }

    [xml] GenerateTimeRequest(){
        $request = $this.xmlDocument.Clone()
        $timeElement = $this.GenerateTimeElement($request)
        $request.DocumentElement.AppendChild($timeElement)
        return $request
    }

    [xml] GenerateWhoamiRequest(){
        $request = $this.xmlDocument.Clone()
        $authElement = $this.GenerateAuthElement($request)
        $request.DocumentElement.AppendChild($authElement)
        $whoamiElement = $this.GenerateWhoamiElement($request)
        $request.DocumentElement.AppendChild($whoamiElement)
        return $request
    }

    [xml] GenerateAuthRequest(){
        $request = $this.xmlDocument.Clone()
        $authElement = $this.GenerateAuthElement($request)
        $request.DocumentElement.AppendChild($authElement)
        return $request
    }

    [xml] GenerateCreateUserRequest([string]$firstName, [string]$lastName, [string]$userEmail, [hashtable]$parameters){
        $request = $this.xmlDocument.Clone()
        $authElement = $this.GenerateAuthElement($request)
        $request.DocumentElement.AppendChild($authElement)
        $createUserElement = $this.GenerateCreateUserElement($request, $firstName, $lastName, $userEmail, $parameters)
        $request.DocumentElement.AppendChild($createUserElement)
        return $request
    }

    [xml] GenerateCreateUserBulkRequest([array]$usersData){
        if ($usersData.Count -ge 1000){
            Write-Host -ForegroundColor Red "Amount of data exceeds limit of 1000 users to be created in one request!"
            throw "Amount of data exceeds limit of 1000 users to be created in one request!"
        }
        $request = $this.xmlDocument.Clone()
        $authElement = $this.GenerateAuthElement($request)
        $request.DocumentElement.AppendChild($authElement)
        foreach($userData in $usersData){
            $createUserElement = $this.GenerateCreateUserElement($request, $userData.firstName, $userData.lastName, $userData.userEmail, $userData.parameters)
            $request.DocumentElement.AppendChild($createUserElement)
        }
        return $request
    }

    [xml] GenerateDeleteUserRequest([array]$userIDs){
        $request = $this.xmlDocument.Clone()
        $authElement = $this.GenerateAuthElement($request)
        $request.DocumentElement.AppendChild($authElement)
        foreach($id in $userIDs){
            $deleteUserElement = $this.GenerateDeleteUserElement($request,$id)
            $request.DocumentElement.AppendChild($deleteUserElement)
        }
        return $request
    }

    [xml] GenerateModifyRequest([string]$type, [string]$id, [hashtable]$dataToUpdate){
        $request = $this.xmlDocument.Clone()
        $authElement = $this.GenerateAuthElement($request)
        $request.DocumentElement.AppendChild($authElement)
        $modifyElement = $this.GenerateModifyElement($request, $type, $id, $dataToUpdate)
        $request.DocumentElement.AppendChild($modifyElement)
        return $request
    }

    [xml] GenerateModifyBulkRequest([array]$modifyRequests){
        $request = $this.xmlDocument.Clone()
        $authElement = $this.GenerateAuthElement($request)
        $request.DocumentElement.AppendChild($authElement)
        foreach($modifyRequest in $modifyRequests){
            $modifyElement = $this.GenerateModifyElement($request, $modifyRequest.type, $modifyRequest.id, $modifyRequest.dataToUpdate)
            $request.DocumentElement.AppendChild($modifyElement)
        }
        return $request
    }
    
    [xml] GenerateAddBulkRequest([array]$addRequests){
        $request = $this.xmlDocument.Clone()
        $authElement = $this.GenerateAuthElement($request)
        $request.DocumentElement.AppendChild($authElement)
        foreach($addRequest in $addRequests){
            $addElement = $this.GenerateAddElement($request, $addRequest.type, $addRequest.dataToAdd)
            $request.DocumentElement.AppendChild($addElement)
        }
        return $request
    }

    [xml] SendRequest([OARequestType] $type, [hashtable]$params){
        $request = $null
        switch($type){
            Time {
                $request = $this.GenerateTimeRequest()
            }
            Read {
                if ($params.limit -and $params.customFields){
                    $request = $this.GenerateReadRequest($params.type, $params.method, $params.queryData, $params.customFields, $params.limit)
                }
                elseif ($params.limit){
                    $request = $this.GenerateReadRequest($params.type, $params.method, $params.queryData, $params.limit)
                }
                elseif ($params.customFields) {
                    $request = $this.GenerateReadRequest($params.type, $params.method, $params.queryData, $params.customFields)
                }
                else {
                    $request = $this.GenerateReadRequest($params.type, $params.method, $params.queryData)
                }
            }
            Whoami {
                $request = $this.GenerateWhoamiRequest()
            }
            Auth {
                $request = $this.GenerateAuthRequest()
            }
            CreateUser {
                $request = $this.GenerateCreateUserRequest($params.firstName, $params.lastName, $params.userEmail, $params.parameters)
            }
            CreateUserBulk {
                $request = $this.GenerateCreateUserBulkRequest($params.usersData)
            }
            DeleteUser {
                $request = $this.GenerateDeleteUserRequest($params.userIDs)
            }
            Modify {
                $request = $this.GenerateModifyRequest($params.type, $params.id, $params.dataToUpdate)
            }
            ModifyBulk {
                $request = $this.GenerateModifyBulkRequest($params.modifyRequests)
            }
            AddBulk {
                Write-Host $params.addRequests | Out-String
                $request = $this.GenerateAddBulkRequest($params.addRequests)
                Write-Host $request.OuterXML
            }
            default {
                $request = $this.xmlDocument.Clone()
            }
        }
        try{
            $response = Invoke-WebRequest $this.OAApiEndpoint -Method POST -Body $request.OuterXml -Headers @{"Content-Type"="application/xml"}
            return [xml]$response.Content
        }
        catch{
            Write-Error "An error occured: $_"
        }
        return $null
    }

    [xml] SendRequest([OARequestType] $type){
        return $this.SendRequest($type, @{})
    }
}