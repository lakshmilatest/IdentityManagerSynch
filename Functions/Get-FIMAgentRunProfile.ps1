function Get-FIMAgentRunProfile
{
    [cmdletbinding()]
    Param(
        [Parameter(
            Mandatory,
            ValueFromPipeline)]
        [Alias("Agent")]
        [string]$Name
    )
    
 BEGIN{
     $f = $Myinvokation.InvokationName
     Write-Verbose -Message "$f - START"
 }
 
 PROCESS
 {
     if(-not $script:MVWebserviceInstance)
     {
         $script:MVWebserviceInstance = Get-MVWebServiceInstance -ErrorAction Stop         
     }
     if($PSBoundParameters.ContainsKey("Name"))
     {
         $name = $name.Replace("'","").Replace('"',"")
     }
     {
        Write-Verbose -Message "$f -  Getting profiles for [$Name]"
        $xmlConfig = $script:MVWebserviceInstance.ExportManagementAgent("$name",$true,$false,[datetime]::Now)
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