<#
.Synopsis
   Get all computers that have the specified software installed
.DESCRIPTION
   Get-CMInstalledSoftware retrieves all computers where the specified software is installed on. 
   The function queries SMS_G_System_INSTALLED_SOFTWARE which contains both 32 and 64 bit software
   installations.

   The following information is collected
   
   Computername,OperatingSystemNameandVersion InstallDate, InstalledLocation, Productname, ProductVersion
   Publisher, UninstallString

.EXAMPLE
   Get-CMInstalledSoftware -ProductName "Client Center for Configuration Manager%"

    Computername                  : Client01
    OperatingSystemNameandVersion : Microsoft Windows NT Workstation 6.1
    Productname                   : Client Center for Configuration Manager 2012
    ProductVersion                : 1.0.1.0
    Publisher                     : Zander Tools
    InstallDate                   : Friday, 1. August 2014 00:00:00
    InstalledLocation             : C:\Program Files\Client Center for Configuration Manager 2012\
    UninstallString               : MsiExec.exe /X{B299EE26-A664-42A2-8D4E-6BF005EB5E75}

.PARAMETER -ProductName
   The Name of the Software. Use % before or after the productname as a wildcard. 
  
#>

Function Get-CMInstalledSoftware
{
    [CmdletBinding()]
    [OutputType([int])]
    Param
    (
        # Software Product Name
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        $ProductName,

        [Parameter(Mandatory=$false,
                   ValueFromPipelineByPropertyName=$true,
                   Position=1)]
        $SiteCode = "SR1",

        [Parameter(Mandatory=$false,
                   ValueFromPipelineByPropertyName=$true,
                   Position=2)]
        $SiteServer = "chrv0300.corp.gwpnet.com"
    )


Begin
{
    [string] $Namespace = "root\SMS\site_$SiteCode"
}


Process
{
    $arsw = Get-WmiObject -Namespace $Namespace -ComputerName $SiteServer -Query "SELECT * FROM SMS_G_System_INSTALLED_SOFTWARE where ProductName like '$ProductName'" # and ProductVersion = '11.0.61030.0'"
    $clientswithsw = @()

    ForEach ($res in $arsw)
    {
        $CMResource = (Get-WmiObject -Namespace $Namespace -ComputerName $SiteServer -Query "SELECT Name, ResourceID,Active, SMSUniqueIdentifier,OperatingSystemNameandVersion FROM SMS_R_SYSTEM WHERE ResourceID = '$($res.ResourceID)'")

        $object = New-Object -TypeName PSObject
        $object | Add-Member -MemberType NoteProperty -Name "Computername" -Value $cmresource.Name
        $object | Add-Member -MemberType NoteProperty -Name "OperatingSystemNameandVersion" -Value $cmresource.OperatingSystemNameandVersion
        $object | Add-Member -MemberType NoteProperty -Name "Productname" -Value $res.ProductName
        $object | Add-Member -MemberType NoteProperty -Name "ProductVersion" -Value $res.ProductVersion
        $object | Add-Member -MemberType NoteProperty -Name "Publisher" -Value $res.Publisher
        $object | Add-Member -MemberType NoteProperty -Name "InstallDate" -Value ([Management.ManagementDateTimeConverter]::ToDateTime($res.InstallDate).DateTime)
        $object | Add-Member -MemberType NoteProperty -Name "InstalledLocation" -Value $res.InstalledLocation
        $object | Add-Member -MemberType NoteProperty -Name "UninstallString" -Value $res.UninstallString

        $clientswithsw += $object
    }
}

End
{
    $clientswithsw
}
}
