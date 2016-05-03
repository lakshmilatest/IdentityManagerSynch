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