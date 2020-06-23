function Get-GPEventByCorrelationID
{
<#
.Synopsis
   Get Group Policy Eventlog entries by Correlation ID
.DESCRIPTION
   This function retrieves Group Policy event log entries filtered by Correlation ID from the specified computer
.EXAMPLE
   Get-GPEventByCorrelationID -Computer TestClient -CorrelationID A2A621EC-44B4-4C56-9BA3-169B88032EFD 
 
TimeCreated                     Id LevelDisplayName Message                                                          
-----------                     -- ---------------- -------                                                          
7/28/2014 5:31:31 PM          5315 Information      Next policy processing for CORP\CHR59104$ will be attempted in...
7/28/2014 5:31:31 PM          8002 Information      Completed policy processing due to network state change for co...
7/28/2014 5:31:31 PM          5016 Information      Completed Audit Policy Configuration Extension Processing in 0...
.......
 
#>
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory=$true,
        ValueFromPipelineByPropertyName=$true,
        HelpMessage="Enter Computername(s)",
        Position=0)]
        [String]$Computer = "localhost",
        # CorrelationID
        [Parameter(Mandatory=$true,
        ValueFromPipelineByPropertyName=$true,
        HelpMessage="Enter CorrelationID",
        Position=0)]
        [string]$CorrelationID
        )
 
    Begin
    {
        $Query = '<QueryList><Query Id="0" Path="Application"><Select Path="Microsoft-Windows-GroupPolicy/Operational">*[System/Correlation/@ActivityID="{CorrelationID}"]</Select></Query></QueryList>'
        $FilterXML = $Query.Replace("CorrelationID",$CorrelationID)
    }
    Process
    {
        $orgCulture = Get-Culture
        [System.Threading.Thread]::CurrentThread.CurrentCulture = New-Object "System.Globalization.CultureInfo" "en-US"
        $gpevents = Get-WinEvent -FilterXml $FilterXML -ComputerName $Computer
        [System.Threading.Thread]::CurrentThread.CurrentCulture = $orgCulture
    }
    End
    {
        [System.Threading.Thread]::CurrentThread.CurrentCulture = New-Object "System.Globalization.CultureInfo" "en-US"
        $gpevents | Format-Table -Wrap -AutoSize -Property TimeCreated, Id, LevelDisplayName, Message
        [System.Threading.Thread]::CurrentThread.CurrentCulture = $orgCulture
     }
}