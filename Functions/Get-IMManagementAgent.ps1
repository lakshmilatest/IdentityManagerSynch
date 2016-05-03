function Get-IMManagementAgent
{
[cmdletbinding()]
[OutputType([CimInstance[]])]
Param(
    [alias("PSComputerName")]
    [string]$ComputerName
    ,
    [Microsoft.Management.Infrastructure.CimSession[]]$Session
)
DynamicParam 
{        
     return Get-DynamicParam -ParameterName "Name" -ParameterAlias "AgentName" -ValueFromPipelineByPropertyName $true -ValueFromPipeline $true
}

BEGIN
{
    $f = $MyInvocation.InvocationName
    Write-Verbose -Message "$f - START"

    $GetCimInstance = @{
        Namespace = "root\MicrosoftIdentityIntegrationServer"
        ClassName = "MIIS_ManagementAgent"
    }
}

PROCESS
{
    [string]$agentName = $PSBoundParameters["Name"]
    
    if($PSBoundParameters.ContainsKey("ComputerName"))
    {
        $null = $GetCimInstance.Add("ComputerName",$ComputerName)
    }
    
    if($PSBoundParameters.ContainsKey("Session"))
    {
        $null = $GetCimInstance.Add("CimSession", $Session)
    }
    
    if($agentName)
    {
        Write-Verbose -Message "$f -  Looking for agent with name [$agentName]"
        $AgentString = $agentName.TrimStart('"').TrimStart("'").TrimEnd('"').TrimEnd("'")
        Get-CimInstance @GetCimInstance -Verbose:$false | Where-Object Name -Like $AgentString     
    }
    else
    {
        Write-Verbose -Message "$f -  Returning all agents"
        Get-CimInstance @GetCimInstance -Verbose:$false
    }     
}

END
{
    Write-Verbose -Message "$f - END"
}

}