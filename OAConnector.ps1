enum OARequestType{
    Read
    CreateUser
    Whoami
    Time
    Auth
}
class OAConnector{
    
    [string] $namespace
    [string] $apiKey
    [hashtable] $OACredentials
    [xml] $xmlDocument
    [string] $OAApiEndpoint = "https://cognifide-ltd.app.sandbox.openair.com/api.pl"

    OAConnector([string]$namespace, [string]$apiKey){
        $this.namespace = $namespace
        $this.apiKey = $apiKey

        Write-Host "Provide OpenAir credentials"
        $company = Read-Host -Prompt "Company"
        $login = Read-Host -Prompt "Login"
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
        $authCompany = $xml.CreateElement("company")
        $authCompany.InnerText = $this.OACredentials.Company
        
        $authUser = $xml.CreateElement("user")
        $authUser.InnerText = $this.OACredentials.User
        
        $authPass = $xml.CreateElement("password")
        $authPass.InnerText = $this.OACredentials.Password
        
        $authLogin = $xml.CreateElement("Login")
        $authLogin.AppendChild($authCompany)
        $authLogin.AppendChild($authUser)
        $authLogin.AppendChild($authPass)

        $auth = $xml.CreateElement("Auth")
        $auth.AppendChild($authLogin)

        return $auth
    }
    
    [System.Xml.XmlElement] GenerateReadElement([xml] $xml, [string]$type, [string]$method, [hashtable]$queryData, [switch]$customFields, [int]$limit){
        $typeElement = $xml.CreateElement($type)
        foreach ($key in $queryData.Keys){
            $queryElement = $xml.CreateElement($key)
            $queryElement.InnerText = $queryData.$key
            $typeElement.AppendChild($queryElement)
        }
        
        $readElement = $xml.CreateElement("Read")
        $readElement.SetAttribute("type","$type")
        $readElement.SetAttribute("method","$method")
        $readElement.SetAttribute("limit","$limit")
        $readElement.SetAttribute("enable_custom",$customFields ? "1" : "0")
        $readElement.AppendChild($typeElement)
        return $readElement
    }

    [System.Xml.XmlElement] GenerateTimeElement($xml){
        $timeElement = $xml.CreateElement("Time")
        $timeElement.InnerText = " "
        return $timeElement
    }
    [System.Xml.XmlElement] GenerateWhoamiElement($xml){
        $timeElement = $xml.CreateElement("Whoami")
        $timeElement.InnerText = " "
        return $timeElement
    }

    [xml] GenerateReadRequest([string]$type, [string]$method, [hashtable]$queryData, [switch]$customFields, [int]$limit){
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

    [xml] GenerateReadRequest([string]$type, [string]$method, [hashtable]$queryData, [switch]$customFields){
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
            Auth{
                $request = $this.GenerateAuthRequest()
            }
            default {
                $request = New-Object -TypeName xml
            }
        }
        try{
            $response = Invoke-WebRequest $this.OAApiEndpoint -Method POST -Body $request.OuterXml -Headers @{"Content-Type"="application/xml"} -Verbose
            return [xml]$response.Content
        }
        catch{
            Write-Error "An error occured: $_"
        }
        return $null
    }
}