function Get-IMAgentRunProfile
{
    [cmdletbinding()]
    Param()
    DynamicParam 
    {        
        return Get-DynamicParam -ParameterName "Name" -Mandatory $true -ParameterAlias "Agent" -ValueFromPipeline $true
    }
    
 BEGIN
 {
     $f = $Myinvokation.InvokationName
     Write-Verbose -Message "$f - START"
 }
 
 PROCESS
 {
    if(-not $script:MVWebserviceInstance)
    {
        Write-Verbose -Message "$f -  Creating a new instance of MVWebService"
        $script:MVWebserviceInstance = Get-MVWebServiceInstance -ErrorAction Stop         
    }
    [String]$Agentname = $PSBoundParameters["Name"]
    $Agentname = $Agentname.TrimStart('"').TrimStart("'").TrimEnd('"').TrimEnd("'")
     
    Write-Verbose -Message "$f -  Getting profiles for [$Agentname]"
    $xmlConfig = $script:MVWebserviceInstance.ExportManagementAgent("$Agentname",$true,$false,[datetime]::Now)
    $xmldoc = [xml]$xmlConfig
    $profiles = $xmldoc.'saved-ma-configuration'.'ma-data'.'ma-run-data'.'run-configuration'.Name
    foreach($name in $profiles)
    {
        "'$name'"
    }        
 }
 
 END{
     Write-Verbose -Message "$f - END"
 }
 
}