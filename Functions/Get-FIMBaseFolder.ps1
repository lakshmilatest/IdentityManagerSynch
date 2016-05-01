function Get-FIMBaseFolder
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