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
