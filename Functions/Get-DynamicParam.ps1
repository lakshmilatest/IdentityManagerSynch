function Get-DynamicParam
{
[OutputType([System.Management.Automation.RuntimeDefinedParameterDictionary])]
[cmdletbinding()]
Param(
    [Parameter(Mandatory)]
    [ValidateSet("Name")]
    $ParameterName
    ,
    [String]$ParameterAlias
    ,
    [bool]$Mandatory = $false
    ,
    [bool]$ValueFromPipelineByPropertyName = $false
    ,
    [bool]$ValueFromPipeline = $false
    ,
    [scriptblock]$ValidateScript
    ,
    $Dictionary
    ,
    [switch]$Mock
)

if($PSBoundParameters.ContainsKey("Dictionary") -eq $false)
{
    $Dictionary = New-Object System.Management.Automation.RuntimeDefinedParameterDictionary
}

$GetCimInstance = @{
        Namespace = "root\MicrosoftIdentityIntegrationServer"
        ClassName = "MIIS_ManagementAgent"
}

$NewDynParam = @{
        Name = "$ParameterName"
        Mandatory = Mandatory
        ValueFromPipelineByPropertyName = $ValueFromPipelineByPropertyName
        ValueFromPipeline = $ValueFromPipeline
        DPDictionary = $Dictionary
}

if($PSBoundParameters.ContainsKey("ParameterAlias"))
{
    $null = $NewDynParam.Add("Alias",$ParameterAlias)
}

if($PSBoundParameters.ContainsKey("ValidateScript"))
{
    $null = $NewDynParam.Add("ValidateScript",$ValidateScript)
}

$agents = $null
$runProfiles = $null

if($Mock)
{
    $agents = @("FIM Service","Active Directory")
    $runProfiles = @("Full Import","Export","Full Synch")
}

switch ($ParameterName) 
{
    "Name" 
    {  
        if(-not $agents)
        {
            $agents = Foreach($agent in (Get-CimInstance @GetCimInstance -Verbose:$false | Select-Object -ExpandProperty Name))
            {
                "'$agent'"
            }      
        }
                
        if($agents)
        {
            $null = $NewDynParam.Add("ValidateSet",$agents)
            New-DynamicParam @NewDynParam -TypeAsString DynParamQuotedString
        }        
    }
    
    "RunProfile" 
    {
        if(-not $runProfiles)
        {
            $runProfiles = Get-IMAagentRunProfile -Name (Get-IMManagementAgent | Select-Object -first 1 -ExpandProperty Name)
        }
        
        If($runProfiles)
        {
            $null = $NewDynParam.Add("ValidateSet",$runProfiles)
            New-DynamicParam @NewDynParam -TypeAsString DynParamQuotedString
        }
    }
    Default {}
}

return $Dictionary
}