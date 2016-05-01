function Get-IMManagementAgent
{
[cmdletbinding()]
[OutputType([CimInstance[]])]
Param(
    [alias("PSComputerName")]
    [string]$ComputerName
    ,
    [Microsoft.Management.Infrastructure.CimSession]$Session
)
DynamicParam {    
    $GetCimInstance = @{
        Namespace = "root\MicrosoftIdentityIntegrationServer"
        ClassName = "MIIS_ManagementAgent"
    }
    $Dictionary = New-Object System.Management.Automation.RuntimeDefinedParameterDictionary
    $Agents = Foreach($agent in (Get-CimInstance @GetCimInstance -Verbose:$false | Select-Object -ExpandProperty Name))
    {
        "'$agent'"
    }
    
    $NewDynParam = @{
        Name = "Name"
        Alias = "AgentName"
        Mandatory = $false
        ValueFromPipelineByPropertyName = $true
        ValueFromPipeline = $true
        DPDictionary = $Dictionary
    }
    if($Agents)
    {
        $null = $NewDynParam.Add("ValidateSet",$Agents)
    }
    New-DynamicParam @NewDynParam -TypeAsString DynParamQuotedString
    return $Dictionary
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
        Get-CimInstance @GetCimInstance -Verbose:$false -CimSession| Where-Object Name -Like $AgentString     
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