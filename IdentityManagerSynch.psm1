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

function Get-IMBaseFolder
{
[cmdletbinding()]
Param()

BEGIN
{
    $f = $MyInvocation.InvocationName
    Write-Verbose -Message "$f - START"
}

PROCESS{}

END 
{
    $fimSyncSvc = Get-CimInstance -ClassName win32_Service | where Name -eq FIMSynchronizationService
    if(-not $fimSyncSvc)
    {
        Write-Verbose -Message "$f -  Unable to find service [FIMSynchronizationService]"
        break
    }
    
    $ServicePath = $fimSyncSvc.PathName.Replace('"','')
    Write-Verbose -Message "$f -  ServicePath is [$ServicePath]"
    $ServicePath | Split-Path | Split-Path
}
}

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

$script:MVWebserviceInstance = $null

function Get-MVWebServiceInstance
{
[cmdletbinding()]
Param(
    [switch]$Force
)

BEGIN 
{
    $f = $MyInvocation.InvocationName
    Write-Verbose -Message "$f - START"
}

PROCESS {}

END 
{
    if($script:MVWebserviceInstance)
    {
        if(-not $Force)
        {
            Write-Verbose -Message "$F -  Force switch not specified, returning object from script scope"
            $script:MVWebserviceInstance
            break
        }
    }
    $fimBasePath = Get-IMBaseFolder
    $fimUIshellPath = "$fimBasePath\UIShell\"
    $propertySheetDll = "PropertySheetBase.dll"
    $assemblyFullPath = "$fimUIshellPath$PropertySheetDll"
    
    Write-Verbose -Message "$f -  Loading assembly [$assemblyFullPath]"
    
    if(-not (Test-Path -Path $assemblyFullPath))
    {
        Write-Error -Message "Unable to find [$assemblyFullPath]" -ErrorAction Stop        
    }
    
    $assemblyPropSheetBase = [System.Reflection.Assembly]::LoadFrom($assemblyFullPath)
    
    Write-Verbose -Message "$f -  Creating MMSWebService instance"
    
    try
    {
        $script:MVWebserviceInstance = $assemblyPropSheetBase.CreateInstance("Microsoft.DirectoryServices.MetadirectoryServices.UI.WebServices.MMSWebService")
        $script:MVWebserviceInstance
    }
    catch
    {
        [System.Exception]$ex = $_.Exception
        Write-Verbose -Message "$f -  Exception $($ex.Message)"
        if($ex.InnerException)
        {
            Write-Verbose -Message "$f-  InnerException $($ex.InnerException.Message)"
        }
    }
    Finally
    {
        Write-Verbose -Message "$f - END"
    }  
}
}

function Invoke-FIMAgentRunProfile
{
<#
.Synopsis
   Invokes a runprofile
.DESCRIPTION
   This cmdlet invokes the CIM-method Activate in class Win32_PowerPlan. See also Get-PowerPlan cmdlet
.EXAMPLE
   Set-PowerPlan -PlanName high*

   This will set the current powerplan to High for the current computer
.EXAMPLE
   Get-Powerplan -PlanName "Power Saver" | Set-PowerPlan

   Will set the powerplan to "Power Saver" for current computer
.EXAMPLE
   Get-Powerplan -PlanName "Power Saver" -ComputerName "Server1","Server2" | Set-PowerPlan

   This will set the current powerpla to "Power Saver" for the computers Server1 and Server2
.EXAMPLE
   Set-PowerPlan -PlanName "Power Saver" -ComputerName "Server1","Server2"

   This will set the current powerpla to "Power Saver" for the computers Server1 and Server2
.NOTES
   Powerplan and performance
.COMPONENT
   Powerplan
.ROLE
   Powerplan
.FUNCTIONALITY
   This cmdlet invokes CIM-methods in the class Win32_PowerPlan
#>
[cmdletbinding(
    SupportsShouldProcess=$true,
    ConfirmImpact='Medium'
)]
[OutputType([PSCustomObject])]
Param(
    [Alias("PSComputerName")]
    [string[]]$ComputerName
    ,
    [int]$DeleteLimit = [int]::MaxValue
    ,
    [int]$UpdateLimit = [int]::MaxValue
    ,
    [int]$TotalLimit = [int]::MaxValue
)
DynamicParam {
    <#
    $Dictionary = New-Object System.Management.Automation.RuntimeDefinedParameterDictionary

    $NewDynParam = @{
        Name = "AgentName"
        Alias = "Name"
        Mandatory = $true
        ValueFromPipelineByPropertyName = $true
        ValueFromPipeline = $true
        DPDictionary = $Dictionary
    }

    $FIMagents = Get-FIMManagementAgent -ErrorAction SilentlyContinue
    if($FIMagents)
    {
        $Agents = @($FIMagents | Select-Object -ExpandProperty Name | ForEach-Object{$_ -replace " ", "-"})
        $null = $NewDynParam.Add("ValidateSet",$Agents)
        $HiddenHash = @{}
        foreach($Agent in $FIMagents)
        {
            $null = $HiddenHash.Add($Agent.Name.Replace(" ","-"), $Agent.Name)
        }
        $PSBoundParameters.Add("Hidden",$HiddenHash)
    }    
    
    New-DynamicParam @NewDynParam #>
    $GetDynamicParam = @{
        ParameterName = "Name"
        ParameterAlias = "AgentName"
        ValueFromPipelineByPropertyName = $true
        ValueFromPipeline = $true
        ValidateScript = {$null = Set-Variable -Name HiddenAgentName -Value $_ -Scope Global ; $true}
    }
    
    $Dictionary = Get-DynamicParam @GetDynamicParam  #-ParameterName "Name" -ParameterAlias "AgentName" -ValueFromPipelineByPropertyName $true -ValueFromPipeline $true
    
    $runProfiles = foreach($profile in (Get-IMAagentRunProfile -Name (Get-IMManagementAgent | Select-Object -first 1 -ExpandProperty Name)))
    {
        $Profile
    }
    
    if(Get-Variable -Scope Global | Where-Object Name -eq "HiddenAgentName")
    {
        $runProfiles = Get-IMAagentRunProfile -Name (Get-Variable -Name HiddenNumberVariable -Scope Global).Value        
    }
    
    $GetDynamicParam.ParameterName = "RunProfile"
    $GetDynamicParam.ParameterAlias = "RunProfileName"
    $GetDynamicParam.ValueFromPipelineByPropertyName = $false
    $GetDynamicParam.ValueFromPipeline = $true
    if($GetDynamicParam.ContainsKey("ValidateScript"))
    {
        $null = $GetDynamicParam.Remove("ValidateScript")
    }
    
    return Get-DynamicParam @GetDynamicParam -Dictionary $Dictionary
    #$DynRunProfiles = @((Get-FIMRunProfile).Keys)
    #New-DynamicParam -Name RunProfile -Type String -ValidateSet $runProfiles -Mandatory -Position 1 -DPDictionary $Dictionary
    
    #$null = out-file -FilePath C:\temp\tore\bound.txt -Append -InputObject "agent name $($PSBoundParameters.AgentName)"
    
}

BEGIN
{
    $f = $MyInvocation.InvocationName
    Write-Verbose -Message "$f - START"
    
    $GetCimInstance = @{
        Namespace = "root\MicrosoftIdentityIntegrationServer"
        ClassName = "MIIS_ManagementAgent"
    }

    if($ComputerName)
    {
        $GetCimInstance.Add("ComputerName",$ComputerName)
    }

    $InvokeCimMethod = @{
        MethodName = "Execute"
    }

    if($WhatIfPreference)
    {
        $InvokeCimMethod.Add("WhatIf",$true)
    }
}

PROCESS
{   
    $AgentName = $PSBoundParameters.AgentName
    $RunProfile =(Get-FIMRunProfile).($PSBoundParameters.RunProfile)
    $AgentNamesHash = $PSBoundParameters.Hidden
    $agentRealName = $AgentNamesHash.$AgentName

    Write-Verbose -Message "$f -  RunProfile = $RunProfile"
    Write-Verbose -Message "$f -  Name       = $AgentName"
    Write-Verbose -Message "$f -  AgentRealName = $agentRealName"
    
    $CimObjectMA = Get-FIMManagementAgent -Name "$agentRealName"

    if(-not $CimObjectMA)
    {
        Write-Warning -Message "Unable to find agent with name '$agentRealName'"
        break
    }   

    foreach($Instance in $CimObjectMA)
    {
        if($pscmdlet.ShouldProcess("$AgentName", "Running profile '$RunProfile'"))
        {
            Write-Verbose -Message "$f -  Invoking runprofile '$RunProfile' for '$agentRealName'"
            $CIMoutput = Invoke-CimMethod -InputObject $Instance @InvokeCimMethod -Arguments @{RunProfileName="$RunProfile"}
        }
        
        Write-Verbose -Message "$f -  Output was $($CIMoutput.ReturnValue)"        

        if($runProfile.Contains("Import") -or $runProfile.Contains("delta"))
        {
            Write-Verbose -Message "$f -  Running import/delta output"
            $outPut = "" | Select-Object -Property IsSuccess, StartTime, Duration, ResultText, ImportChanges, ImportAdds, ImportDeletes, TotalChanges, MA, DeleteLimit, UpdateLimit, RunProfile
            $outPut.UpdateLimit = $UpdateLimit
            $outPut.DeleteLimit = $DeleteLimit 
            $outPut.RunProfile = $RunProfile
            $outPut.MA = $AgentName
            $output.ResultText = $CIMoutput.ReturnValue
            $outPut.IsSuccess = $false
            if($outPut.ResultText -eq "success")
            {
                $outPut.IsSuccess = $true
            }
            $outPut.StartTime = [datetime](Invoke-FIMManagementAgentMethod -AgentName $AgentName -MethodName RunStartTime -ErrorAction SilentlyContinue | Select-Object -ExpandProperty ReturnValue)
            $endtime = [datetime](Invoke-FIMManagementAgentMethod -AgentName $AgentName -MethodName RunEndTime -ErrorAction SilentlyContinue | Select-Object -ExpandProperty ReturnValue)
            $outPut.Duration = New-TimeSpan -Start $output.StartTime -End $endtime
            $outPut.ImportAdds = [int](Invoke-FIMManagementAgentMethod -AgentName $AgentName -MethodName "NumImportAdd" -ErrorAction SilentlyContinue | Select-Object -ExpandProperty ReturnValue)
            $outPut.ImportChanges = [int](Invoke-FIMManagementAgentMethod -AgentName $AgentName -MethodName NumImportUpdate -ErrorAction SilentlyContinue | Select-Object -ExpandProperty ReturnValue)
            $outPut.ImportDeletes = [int](Invoke-FIMManagementAgentMethod -AgentName $AgentName -MethodName NumImportDelete -ErrorAction SilentlyContinue | Select-Object -ExpandProperty ReturnValue)
            $output.TotalChanges = [int]$outPut.ImportAdds + [int]$outPut.ImportChanges + [int]$outPut.ImportDeletes
            $outPut

            if($outPut.ImportDeletes -gt $DeleteLimit)
            {
                Write-Error -Message "DeleteLimit($DeleteLimit) reached, stopping because number of importdeletes is $($outPut.ImportDeletes)"                
            }

            if($outPut.ImportChanges -gt $UpdateLimit)
            {
                Write-Error -Message "UpdateLimit($UpdateLimit) reached, stopping because number of ImportChanges is $($outPut.ImportChanges)"
            }

            if($outPut.TotalChanges -gt $TotalLimit)
            {
                Write-Error -Message "TotalLimit reached, stopping, stopping because number of TotalChanges is $($outPut.TotalChanges)"
            }
        }
        else
        {
            Write-Verbose -Message "$f -  Running export output"
            $outPut = "" | Select-Object -Property IsSuccess, StartTime, Duration, ResultText, ExportChanges, ExportAdds, ExportDeletes, TotalChanges, MA, DeleteLimit, UpdateLimit, RunProfile
            $outPut.UpdateLimit = $UpdateLimit
            $outPut.DeleteLimit = $DeleteLimit
            $outPut.RunProfile = $RunProfile
            $outPut.MA = $AgentName
            $output.ResultText = $CIMoutput.ReturnValue
            $outPut.IsSuccess = $false
            if($outPut.ResultText -eq "success")
            {
                $outPut.IsSuccess = $true
            }
            $outPut.StartTime = [datetime](Invoke-FIMManagementAgentMethod -AgentName $AgentName -MethodName RunStartTime -ErrorAction SilentlyContinue | Select-Object -ExpandProperty ReturnValue)
            $endtime = [datetime](Invoke-FIMManagementAgentMethod -AgentName $AgentName -MethodName RunEndTime -ErrorAction SilentlyContinue | Select-Object -ExpandProperty ReturnValue)
            $outPut.Duration = New-TimeSpan -Start $output.StartTime -End $endtime
            $outPut.ExportAdds = [int](Invoke-FIMManagementAgentMethod -AgentName $AgentName -MethodName NumExportAdd -ErrorAction SilentlyContinue | Select-Object -ExpandProperty ReturnValue)
            $outPut.ExportChanges = [int](Invoke-FIMManagementAgentMethod -AgentName $AgentName -MethodName NumExportUpdate -ErrorAction SilentlyContinue | Select-Object -ExpandProperty ReturnValue)
            $outPut.ExportDeletes = [int](Invoke-FIMManagementAgentMethod -AgentName $AgentName -MethodName NumExportDelete -ErrorAction SilentlyContinue | Select-Object -ExpandProperty ReturnValue)
            $output.TotalChanges = [int]$outPut.ExportAdds + [int]$outPut.ExportChanges + [int]$outPut.ExportDeletes
            $outPut
        }
    }    
}

END
{
    Write-Verbose -Message "$f - END"
}

}


function New-DynamicParam 
{
<#
    .SYNOPSIS
        Helper function to simplify creating dynamic parameters
    
    .DESCRIPTION
        Helper function to simplify creating dynamic parameters

        Example use cases:
            Include parameters only if your environment dictates it
            Include parameters depending on the value of a user-specified parameter
            Provide tab completion and intellisense for parameters, depending on the environment

        Please keep in mind that all dynamic parameters you create will not have corresponding variables created.
           One of the examples illustrates a generic method for populating appropriate variables from dynamic parameters
           Alternatively, manually reference $PSBoundParameters for the dynamic parameter value

    .NOTES
        Credit to http://jrich523.wordpress.com/2013/05/30/powershell-simple-way-to-add-dynamic-parameters-to-advanced-function/
            Added logic to make option set optional
            Added logic to add RuntimeDefinedParameter to existing DPDictionary
            Added a little comment based help
        Credit to https://github.com/RamblingCookieMonster/PowerShell/blob/master/New-DynamicParam.ps1

        Credit to BM for alias and type parameters and their handling

    .PARAMETER Name
        Name of the dynamic parameter

    .PARAMETER Type
        Type for the dynamic parameter.  Default is string

    .PARAMETER Alias
        If specified, one or more aliases to assign to the dynamic parameter

    .PARAMETER ValidateSet
        If specified, set the ValidateSet attribute of this dynamic parameter

    .PARAMETER Mandatory
        If specified, set the Mandatory attribute for this dynamic parameter

    .PARAMETER ParameterSetName
        If specified, set the ParameterSet attribute for this dynamic parameter

    .PARAMETER Position
        If specified, set the Position attribute for this dynamic parameter

    .PARAMETER ValueFromPipelineByPropertyName
        If specified, set the ValueFromPipelineByPropertyName attribute for this dynamic parameter

    .PARAMETER HelpMessage
        If specified, set the HelpMessage for this dynamic parameter
    
    .PARAMETER DPDictionary
        If specified, add resulting RuntimeDefinedParameter to an existing RuntimeDefinedParameterDictionary (appropriate for multiple dynamic parameters)
        If not specified, create and return a RuntimeDefinedParameterDictionary (appropriate for a single dynamic parameter)

        See final example for illustration

    .EXAMPLE
        
        function Show-Free
        {
            [CmdletBinding()]
            Param()
            DynamicParam {
                $options = @( gwmi win32_volume | %{$_.driveletter} | sort )
                New-DynamicParam -Name Drive -ValidateSet $options -Position 0 -Mandatory
            }
            begin{
                #have to manually populate
                $drive = $PSBoundParameters.drive
            }
            process{
                $vol = gwmi win32_volume -Filter "driveletter='$drive'"
                "{0:N2}% free on {1}" -f ($vol.Capacity / $vol.FreeSpace),$drive
            }
        } #Show-Free

        Show-Free -Drive <tab>

    # This example illustrates the use of New-DynamicParam to create a single dynamic parameter
    # The Drive parameter ValidateSet populates with all available volumes on the computer for handy tab completion / intellisense

    .EXAMPLE

    # I found many cases where I needed to add more than one dynamic parameter
    # The DPDictionary parameter lets you specify an existing dictionary
    # The block of code in the Begin block loops through bound parameters and defines variables if they don't exist

        Function Test-DynPar{
            [cmdletbinding()]
            param(
                [string[]]$x = $Null
            )
            DynamicParam
            {
                #Create the RuntimeDefinedParameterDictionary
                $Dictionary = New-Object System.Management.Automation.RuntimeDefinedParameterDictionary
        
                New-DynamicParam -Name AlwaysParam -ValidateSet @( gwmi win32_volume | %{$_.driveletter} | sort ) -DPDictionary $Dictionary

                #Add dynamic parameters to $dictionary
                if($x -eq 1)
                {
                    New-DynamicParam -Name X1Param1 -ValidateSet 1,2 -mandatory -DPDictionary $Dictionary
                    New-DynamicParam -Name X1Param2 -DPDictionary $Dictionary
                    New-DynamicParam -Name X3Param3 -DPDictionary $Dictionary -Type DateTime
                }
                else
                {
                    New-DynamicParam -Name OtherParam1 -Mandatory -DPDictionary $Dictionary
                    New-DynamicParam -Name OtherParam2 -DPDictionary $Dictionary
                    New-DynamicParam -Name OtherParam3 -DPDictionary $Dictionary -Type DateTime
                }
        
                #return RuntimeDefinedParameterDictionary
                $Dictionary
            }
            Begin
            {
                #This standard block of code loops through bound parameters...
                #If no corresponding variable exists, one is created
                    #Get common parameters, pick out bound parameters not in that set
                    Function _temp { [cmdletbinding()] param() }
                    $BoundKeys = $PSBoundParameters.keys | Where-Object { (get-command _temp | select -ExpandProperty parameters).Keys -notcontains $_}
                    foreach($param in $BoundKeys)
                    {
                        if (-not ( Get-Variable -name $param -scope 0 -ErrorAction SilentlyContinue ) )
                        {
                            New-Variable -Name $Param -Value $PSBoundParameters.$param
                            Write-Verbose "Adding variable for dynamic parameter '$param' with value '$($PSBoundParameters.$param)'"
                        }
                    }

                #Appropriate variables should now be defined and accessible
                    Get-Variable -scope 0
            }
        }

    # This example illustrates the creation of many dynamic parameters using New-DynamicParam
        # You must create a RuntimeDefinedParameterDictionary object ($dictionary here)
        # To each New-DynamicParam call, add the -DPDictionary parameter pointing to this RuntimeDefinedParameterDictionary
        # At the end of the DynamicParam block, return the RuntimeDefinedParameterDictionary
        # Initialize all bound parameters using the provided block or similar code

    .FUNCTIONALITY
        PowerShell Language

#>
param(    
    [string]$Name
    ,
    [System.Type]$Type = [string]
    ,
    [string]$TypeAsString
    ,
    [string[]]$Alias = @()
    ,
    [string[]]$ValidateSet
    ,    
    [scriptblock]$validateScript
    ,
    [switch]$Mandatory
    ,    
    [string]$ParameterSetName = "__AllParameterSets"
    ,    
    [int]$Position
    ,    
    [switch]$ValueFromPipelineByPropertyName
    ,
    [switch]$ValueFromPipeline
    ,    
    [string]$HelpMessage
    ,
    [validatescript({
        if(-not ( $_ -is [System.Management.Automation.RuntimeDefinedParameterDictionary] -or -not $_) )
        {
            Throw "DPDictionary must be a System.Management.Automation.RuntimeDefinedParameterDictionary object, or not exist"
        }
        $True
    })]
    $DPDictionary = $false
 
)
    Add-Type @"
    public class DynParamQuotedString {
 
        public DynParamQuotedString(string quotedString) : this(quotedString, "'") {}
        public DynParamQuotedString(string quotedString, string quoteCharacter) {
            OriginalString = quotedString;
            _quoteCharacter = quoteCharacter;
        }

        public string OriginalString { get; set; }
        string _quoteCharacter;

        public override string ToString() {
            if (OriginalString.Contains(" ")) {
                return string.Format("{1}{0}{1}", OriginalString, _quoteCharacter);
            }
            else {
                return OriginalString;
            }
        }
    }
"@ 
    if($PSBoundParameters.ContainsKey("TypeAsString"))
    {
        $type = [System.Type]$TypeAsString
    }
    #Create attribute object, add attributes, add to collection   
        $ParamAttr = New-Object System.Management.Automation.ParameterAttribute
        $ParamAttr.ParameterSetName = $ParameterSetName
        if($mandatory)
        {
            $ParamAttr.Mandatory = $True
        }
        if($Position -ne $null)
        {
            $ParamAttr.Position=$Position
        }
        if($ValueFromPipelineByPropertyName)
        {
            $ParamAttr.ValueFromPipelineByPropertyName = $True            
        }
        if($ValueFromPipeline)
        {
            $ParamAttr.ValueFromPipeline = $True
        }
        if($HelpMessage)
        {
            $ParamAttr.HelpMessage = $HelpMessage
        }
 
        $AttributeCollection = New-Object 'Collections.ObjectModel.Collection[System.Attribute]'
        $AttributeCollection.Add($ParamAttr)
    
    #param validation set if specified
        if($ValidateSet)
        {
            $ParamOptions = New-Object System.Management.Automation.ValidateSetAttribute -ArgumentList $ValidateSet
            $AttributeCollection.Add($ParamOptions)
        }
        
        if($validateScript)
        {
            $paramScript = New-Object -TypeName System.Management.Automation.ValidateScriptAttribute -ArgumentList $validateScript
            $AttributeCollection.Add($paramScript)
        }

    #Aliases if specified
        if($Alias.count -gt 0) {
            $ParamAlias = New-Object System.Management.Automation.AliasAttribute -ArgumentList $Alias
            $AttributeCollection.Add($ParamAlias)
        }

 
    #Create the dynamic parameter
        $Parameter = New-Object -TypeName System.Management.Automation.RuntimeDefinedParameter -ArgumentList @($Name, $Type, $AttributeCollection)
    
    #Add the dynamic parameter to an existing dynamic parameter dictionary, or create the dictionary and add it
        if($DPDictionary)
        {
            $DPDictionary.Add($Name, $Parameter)
        }
        else
        {
            $Dictionary = New-Object System.Management.Automation.RuntimeDefinedParameterDictionary
            $Dictionary.Add($Name, $Parameter)
            $Dictionary
        }
}


