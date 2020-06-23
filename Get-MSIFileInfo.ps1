function Get-MSIFileInfo()
{
<#
.SYNOPSIS
    This function retrieves File information from Windows Installer (MSI) files
.DESCRIPTION
    This fucntion uses the Windows Instaler COM object to retrieve file information from the Windows Installer
    database. The script can be used for analyzing content of Windows Installer databases without the need of
    having to initiate an actual install. 

    The function retrieves the following information for each file referenced in the Windows Installer Database

    MSIFileFullname   : C:\TEMP\cmcollctr_1.0.0.11.msi
    MSIProductName    : Collection Commander for Configuration Manager
    MSIProductVersion : 1.0.0.11
    Manufacturer      : Zander Tools
    MSIProductCode    : {38253945-0CBC-4A05-BCAF-CE979EA4AAEF}
    File              : sccmclictr.automation.dll
    Component         : sccmclictr.automation.dll
    FileName          : sccmcl~1.dll|sccmclictr.automation.dll
    FileSize          : 435576
    Version           : 0.0.0.55
    Directory         : APPDIR
    Directory_Parent  : TARGETDIR
    DefaultDir        : APPDIR:.
    TargetPath        : C:\Program Files (x86)\Collection Commander for Configuration Manager

.EXAMPLE
    Get-MSIFileInfo -Installerfile "C:\TEMP\cmcollctr_1.0.0.11.msi"

    This command processes the specified file(s) and stores the file information results into a csv file
    located in the same folder as the script. 

.EXAMPLE
    Get-MSIFileInfo -IstallerFolder "C:\TEMP\MSI"    

    This command processes all MSI files stored under c:\temp\msi and stores the file information results 
    into a csv file located in the same folder as the script. 

.EXAMPLE
    Get-MSIFileInfo -Installerfile "C:\TEMP\cmcollctr_1.0.0.11.msi" -OutPutFile "C:\TEMP\msifileresults.csv"
    
    This command processes the specified file and stores the file informatoin results into the spevified
    output file.    

.PARAMETER InstallerFile
    The path to a Windows Installer file. 

.PARAMETER InstallerFolder
    The path of the folder containing Wwindows Installer files

.PARAMETER OutPutFile (Optional)
    The path to the output file. If no Output file is specified the results are stored into a csv file
    located within the script file folder. 

.LINKS
    MSDN Windows Installer
    http://msdn.microsoft.com/en-us/library/aa369432(v=vs.85).aspx
    http://msdn.microsoft.com/en-us/library/aa371653(v=vs.85).aspx
    http://stackoverflow.com/questions/23903254/advanced-installer-powershell-script-set-property
.NOTES
    Version 1.1.0 by Alex Verboon

    Credits 
    weberik's post at stackoverflow for the vb script code and Claude Henchoz for the conversion into PS
    https://stackoverflow.com/questions/17543132/how-can-i-resolve-msi-paths-in-vbscript#new-answer
    Adam Bertram's Get-MSIProperties script got me started with how to read content from MSI files 
    http://gallery.technet.microsoft.com/scriptcenter/Get-all-native-properties-e4e19180
    Trevor Sullivan's memory usage snippet
    http://trevorsullivan.net/2012/01/23/powershell-prompt-function-to-monitor-memory-usage/
#>
[CmdletBinding()]
param (
[Parameter(ParameterSetName="File",
    Mandatory=$True,
    ValueFromPipeline=$True,
    ValueFromPipelineByPropertyName=$True,
    Position=0,
    HelpMessage='What is the path of the Windows Installer file to query?')]
    [String[]]$InstallerFile, 
[Parameter(ParameterSetName="Directory",
    Mandatory=$True,
    ValueFromPipeline=$True,
    ValueFromPipelineByPropertyName=$True,
    Position=0,
    HelpMessage='What is the path containing the Windows Installer files to query?')]
    [string]$InstallerFolder,
[Parameter(Mandatory=$false,
    Position=1,
    HelpMessage='What is the path of the output file')]
    [String]$OutPutFile
)
      
begin
{
    # Check what parameter was used File or Folder
    switch ($PsCmdlet.ParameterSetName) 
    {
    "File" {
            # Check if the Windows Installer file(s) exists, exit if not
            ForEach ($checkfile in $InstallerFile)
                {
                    if (!(Test-Path -literalpath $checkfile)) 
                    { throw "File '{0}' does not exist" -f $checkfile}
                }
            }
    "Directory"{
            # Check if the provided Directory exists, exit if not
            if (!(Test-Path -Path $InstallerFolder))   
                {
                    throw "Folder '{0}' does not exist" -f $InstallerFolder
                }
            Else
                {
                    # When a folder is provided, retrieve all Windows installer files
                    # in the folder and subfolders 
                    $InstallerFile = (Get-ChildItem -Path "$InstallerFolder\*.MSI" -Recurse -File).FullName
                }
            }
    } # end switch


    # Check if an output file was specified
    if ($PSBoundParameters.ContainsKey("OutPutFile"))
        {
            # no Output file parameter specified so we use the default output defined below
            $msifileinfo_output = $OutPutFile
        }
    Else
        {
            # construct the results output file
            $timestamp = $((get-date).tostring("MMddyyyyHHmmss"))
            $filename =  "MSI_FileInfo" + "_" + $timestamp + ".txt"
            $msifileinfo_output = $PSScriptRoot + "\" + "$filename"
        }


    # Windows Installer COM object
    $com_object = New-Object -com WindowsInstaller.Installer

    # construct the temp powershell script file name 
    $tmpfile =   [guid]::NewGuid().Guid + ".ps1"
    $tmpfld = "$env:Temp\"
    $tmpps = $tmpfld + $tmpfile

    # $launchmsiscript contains the powershell code used to run the Windows Installer CostInitialize and
    # Costfinalize actions
    # http://msdn.microsoft.com/en-us/library/aa368050(v=vs.85).aspx
    # http://msdn.microsoft.com/en-us/library/aa368048(v=vs.85).aspx
    # we have to launch a separate powershell process as otherwise the Windows Installer process remains open
    # the below code is stored into a temporary ps1 file that is then launched from the below process.   

$launchmsiscript = @"
<#
.Synopsis
   This function invokes the Widnows Installer CostFinalize action
.DESCRIPTION
   This function invokes the Widnows Installer CostFinalize action
   and then provides the filename, component and resolved path of the path as output 
.EXAMPLE
    Invoke-MSICostFinalize -FileName "C:\TEMP\cmcollctr_1.0.0.11.msi"
.NOTES
    Credits
    weberik's post at stackoverflow for the vb script code to run the MSI costfinalize action
    https://stackoverflow.com/questions/17543132/how-can-i-resolve-msi-paths-in-vbscript#new-answer
    Claude Henchoz for converting the above referenced vb script code into powershell
#>
function Invoke-MSICostFinalize
{
    [CmdletBinding()]
    Param
    (
        # Param1 help description
        [Parameter(Mandatory=`$true,
                   ValueFromPipelineByPropertyName=`$true,
                   Position=0)]
        [String]`$FileName 
    )

Begin
{
        # Check if the Installer File exists
        if (!(Test-Path -literalpath `$FileName)) 
        { throw "File '{0}' does not exist" -f `$FileName}
}

Process{

    # Installer init, OpenDatabase, OpenPackage
    `$Installer = New-Object -ComObject "WindowsInstaller.Installer"
    `$DB = `$Installer.GetType().InvokeMember("OpenDatabase","InvokeMethod",`$Null,`$Installer,@(`$FileName, 0))
    `$Session = `$Installer.GetType().InvokeMember("OpenPackage","InvokeMethod",`$Null,`$Installer,@(`$DB, 0))

    # CostInitialize, CostFinalize
    `$Session.GetType().InvokeMember("DoAction","InvokeMethod",`$Null,`$Session,@("CostInitialize", 0))
    `$Session.GetType().InvokeMember("DoAction","InvokeMethod",`$Null,`$Session,@("CostFinalize", 0))

    # Add Query
    `$View = `$DB.GetType().InvokeMember("OpenView","InvokeMethod",`$Null,`$DB,
        @("SELECT File, Directory_, FileName, Component_, Component FROM File,Component WHERE Component=Component_ ORDER BY Directory_", 0)
    )

    # Execute Query
    `$View.GetType().InvokeMember("Execute","InvokeMethod",`$Null,`$View,@(0))

    # Fetch 1st record
    `$Record = `$View.GetType().InvokeMember("Fetch","InvokeMethod",`$Null,`$View,@(0))

    do {
        # Fill variables
        `$File = `$Record.GetType().InvokeMember("StringData","GetProperty",`$Null,`$Record,@(1,0))
        `$DirectoryName = `$Record.GetType().InvokeMember("StringData","GetProperty",`$Null,`$Record,@(2,0))
        `$FileName = `$Record.GetType().InvokeMember("StringData","GetProperty",`$Null,`$Record,@(3,0))
        `$Components = `$Record.GetType().InvokeMember("StringData","GetProperty",`$Null,`$Record,@(4,0))
    
        # Split filename
        try {
            `$FileName = `$FileName.Split("|")[1]
        } catch {}

        # Resolve Directory
        `$ResolvedDirectory = `$Session.GetType().InvokeMember("TargetPath","GetProperty",`$Null,`$Session,@(`$DirectoryName, 0))
    
        # Output

       `$outline =  `$ResolvedDirectory + `$FileName + "," + `$Components

        write-output `$outline

        # Get next record
        `$Record = `$View.GetType().InvokeMember("Fetch","InvokeMethod",`$Null,`$View,@(0))
    } while (`$Record)
}

End{}
} # end function


If ([string]::IsNullOrEmpty(`$args) -ne `$true)
{
    Invoke-MSICostFinalize -FileName "`$args"
}
Else
{
    "Input Parameter -FileName is missing"
}
"@

#write the temporary powershell script file
$launchmsiscript | Out-File -FilePath $tmpps -NoClobber -Encoding ascii
}


process
{
    $ParentID = 1
    $swcount = $InstallerFile.count
    $si=1

    ForEach ($msifile in $InstallerFile)
    {
        $FilePath = [IO.FileInfo[]]$msifile
        Write-Progress -id $ParentID -Activity "Processing  $si / $swcount" -Status "$msifile" -PercentComplete (($si / $swcount) * 100)

        # construct the temp resolved path output file 
        $tmpfile =   [guid]::NewGuid().Guid + ".txt"
        $tmpfld = "$env:Temp\"
        $tmprpaths = $tmpfld + $tmpfile

        # launch the temp powershell script to resolve the MSI paths
        Write-Progress -ParentId $ParentID -Activity "Initialization" -Status "Resolving paths" -PercentComplete 0
        $resolvedpath = powershell.exe -NoLogo -file "$tmpps" "$FilePath" 
        Write-Progress -ParentId $ParentID -Activity "Initialization" -Status "Resolving paths completed" -PercentComplete 100

        # store the results into a tmp file
        $resolvedpath | Out-File -FilePath $tmprpaths -Encoding ascii

        # bulid the resolved paths table
        $in = 0
        $resolvedpathdetails = @()
        while ($in -lt $resolvedpath.Count)
        {
            $fullpath = $resolvedpath.item($in).split(",")[0]
            $component = $resolvedpath.item($in).split(",")[1]
            $in = $in + 1
            $rpath = New-Object -TypeName psobject
            $rpath | Add-Member -MemberType NoteProperty -Name "FullPath" -Value $fullpath
            $rpath | Add-Member -MemberType NoteProperty -Name "Component" -Value $component
            $resolvedpathdetails += $rpath
        }

        try 
        {
            # -------------------------------------------------------------------------------------------------#
            # Open Database 1
            # -------------------------------------------------------------------------------------------------#
            Write-Progress -ParentId $ParentID -Activity "Initialization" -Status "Opening Database" -PercentComplete 0

            $database = $com_object.GetType().InvokeMember(
                "OpenDatabase",
                "InvokeMethod",
                $Null,
                $com_object,
                @($FilePath.FullName, 0)
            )

            # -------------------------------------------------------------------------------------------------#
            # MSI Property Information 2
            # -------------------------------------------------------------------------------------------------#
            Write-Progress -ParentId $ParentID -Activity "Initialization" -Status "Gathering Property Information" -PercentComplete ((5/100*1) *100)
            $Propertyquery = "SELECT * FROM Property"
            $PropertyView = $database.GetType().InvokeMember(
                    "OpenView",
                    "InvokeMethod",
                    $Null,
                    $database,
                    ($Propertyquery)
            )

            $PropertyView.GetType().InvokeMember("Execute", "InvokeMethod", $Null, $PropertyView, $Null)

            $Propertyrecord = $PropertyView.GetType().InvokeMember(
                    "Fetch",
                    "InvokeMethod",
                    $Null,
                    $PropertyView,
                    $Null
            )

            $properties_table = @{}
            while ($Propertyrecord -ne $null)
            {
            $Prop =  $Propertyrecord.GetType().InvokeMember("StringData", "GetProperty", $Null, $Propertyrecord, 1)
            $value = $Propertyrecord.GetType().InvokeMember("StringData", "GetProperty", $Null, $Propertyrecord, 2)
            $properties_table.Add("$($prop)","$($value)")

            $Propertyrecord = $PropertyView.GetType().InvokeMember(
                    "Fetch",
                    "InvokeMethod",
                    $Null,
                    $PropertyView,
                    $Null
                )
            }

            # -------------------------------------------------------------------------------------------------#
            # MSI Directory Information 3
            # -------------------------------------------------------------------------------------------------#
            Write-Progress -ParentId $ParentID -Activity "Initialization" -Status "Gathering Directory Information" -PercentComplete ((5/100*2) *100)
            $dirquery = "SELECT * FROM Directory"
            $dirView = $database.GetType().InvokeMember(
                    "OpenView",
                    "InvokeMethod",
                    $Null,
                    $database,
                    ($dirquery)
            )

            $dirView.GetType().InvokeMember("Execute", "InvokeMethod", $Null, $dirView, $Null)
            $dirrecord = $dirView.GetType().InvokeMember(
                    "Fetch",
                    "InvokeMethod",
                    $Null,
                    $dirView,
                    $Null
            )

            
            $directory_table = @{}
            while ($dirrecord -ne $null)
            {
            $Directory =  $dirrecord.GetType().InvokeMember("StringData", "GetProperty", $Null, $dirrecord, 1)
            $Directory_Parent = $dirrecord.GetType().InvokeMember("StringData", "GetProperty", $Null, $dirrecord, 2)
            $DefaultDir = $dirrecord.GetType().InvokeMember("StringData", "GetProperty", $Null, $dirrecord, 3)
            $directory_table.Add("$Directory",("$Directory_Parent","$DefaultDir"))
            $dirrecord = $dirView.GetType().InvokeMember(
                "Fetch",
                "InvokeMethod",
                $Null,
                $dirView,
                $Null
                )
            }

            # -------------------------------------------------------------------------------------------------#
            # MSI File Information 4
            # -------------------------------------------------------------------------------------------------#
            Write-Progress -ParentId $ParentID -Activity "Initialization" -Status "Gathering File Information" -PercentComplete ((5/100*3) *100)
            $filequery = "SELECT File,Component_,FileName,FileSize,Version FROM File"
            $fileView = $database.GetType().InvokeMember(
                    "OpenView",
                    "InvokeMethod",
                    $Null,
                    $database,
                    ($filequery)
            )

            $fileView.GetType().InvokeMember("Execute", "InvokeMethod", $Null, $fileView, $Null)
            $filerecord = $fileView.GetType().InvokeMember(
                    "Fetch",
                    "InvokeMethod",
                    $Null,
                    $fileView,
                    $Null
            )
            
            $files_table = @{}
            while ($filerecord -ne $null)
            {
            $file =  $filerecord.GetType().InvokeMember("StringData", "GetProperty", $Null, $filerecord, 1)
            $component = $filerecord.GetType().InvokeMember("StringData", "GetProperty", $Null, $filerecord, 2)
            $filename = $filerecord.GetType().InvokeMember("StringData", "GetProperty", $Null, $filerecord, 3)
            $filesize = $filerecord.GetType().InvokeMember("StringData", "GetProperty", $Null, $filerecord, 4)
            $fileversion = $filerecord.GetType().InvokeMember("StringData", "GetProperty", $Null, $filerecord, 5)

            $files_table.Add("$File",("$Component","$filename","$filesize","$fileversion"))

            $filerecord = $fileView.GetType().InvokeMember(
                    "Fetch",
                    "InvokeMethod",
                    $Null,
                    $fileView,
                    $Null
                )
            }

            # -------------------------------------------------------------------------------------------------#
            # MSI Component Information 5
            # -------------------------------------------------------------------------------------------------#
            Write-Progress -ParentId $ParentID -Activity "Initialization" -Status "Gathering Component Information" -PercentComplete ((5/100*4) *100)
            $compquery = "SELECT * FROM Component"
            $compView = $database.GetType().InvokeMember(
                    "OpenView",
                    "InvokeMethod",
                    $Null,
                    $database,
                    ($compquery)
            )

            $compView.GetType().InvokeMember("Execute", "InvokeMethod", $Null, $compView, $Null)
            $comprecord = $compView.GetType().InvokeMember(
                    "Fetch",
                    "InvokeMethod",
                    $Null,
                    $compView,
                    $Null
            )

            $components_table = @{}
            while ($comprecord -ne $null)
            {
            $component =  $comprecord.GetType().InvokeMember("StringData", "GetProperty", $Null, $comprecord, 1)
            $componentid = $comprecord.GetType().InvokeMember("StringData", "GetProperty", $Null, $comprecord, 2)
            $directory = $comprecord.GetType().InvokeMember("StringData", "GetProperty", $Null, $comprecord, 3)
            $keypath = $comprecord.GetType().InvokeMember("StringData", "GetProperty", $Null, $comprecord, 4)
            $components_table.Add("$component",("$componentid","$directory","$keypath"))

            $comprecord = $compView.GetType().InvokeMember(
                    "Fetch",
                    "InvokeMethod",
                    $Null,
                    $compView,
                    $Null
                )
            }

            Write-Progress -ParentId $ParentID -Activity "Initialization" -Status "Completed Gathering Information" -PercentComplete ((5/100*5) *100)

            # ------------------------------------------------------------------------------------------#
            # Putting it together
            # ------------------------------------------------------------------------------------------#
            # Get MSI Property data
            $MSIFileFullname = ($FilePath.FullName)
            $MSIProductName = $properties_table["ProductName"]
            $MSIProductVersion = $properties_table["ProductVersion"]
            $Manufacturer = $properties_table["Manufacturer"]
            $MSIProductCode = $properties_table["ProductCode"]

            # load the tmp resolved paths dat into a variable
            $rpf = Get-Content -Path $tmprpaths


            # -------------------------------------------------------------------------------------------------#
            # Create the dataset
            # -------------------------------------------------------------------------------------------------#
            $msi_fileprops = @()
            $fc=1
            $sw = [System.Diagnostics.Stopwatch]::StartNew()

            $totalfiles = $files_table.count
            ForEach($item in $files_table.GetEnumerator()){
                $msifileinfo = New-Object -TypeName psobject
                # MSI Properties data
                $msifileinfo | Add-Member -MemberType NoteProperty -Name "MSIFileFullname" -Value $MSIFileFullname
                $msifileinfo | Add-Member -MemberType NoteProperty -Name "MSIProductName" -Value $MSIProductName
                $msifileinfo | Add-Member -MemberType NoteProperty -Name "MSIProductVersion" -Value $MSIProductVersion
                $msifileinfo | Add-Member -MemberType NoteProperty -Name "Manufacturer" -Value $Manufacturer
                $msifileinfo | Add-Member -MemberType NoteProperty -Name "MSIProductCode" -Value $MSIProductCode
                # MSI File data
                $msifileinfo | Add-Member -MemberType NoteProperty -Name "File" -Value $item.Name
                $msifileinfo | Add-Member -MemberType NoteProperty -Name "Component" -Value $item.Value[0]
                $msifileinfo | Add-Member -MemberType NoteProperty -Name "FileName" -Value $item.Value[1]
                $msifileinfo | Add-Member -MemberType NoteProperty -Name "FileSize" -Value $item.Value[2]
                $msifileinfo | Add-Member -MemberType NoteProperty -Name "Version" -Value $item.Value[3]
                
                 # Get the Directory from the components table
                $msifileinfo | Add-Member -MemberType NoteProperty -Name "Directory" -Value ($cd = $components_table["$($msifileinfo.Component)"][1])
                # Get the Parent Directory from the directory table
                $msifileinfo | Add-Member -MemberType NoteProperty -Name "Directory_Parent" ($dp = $directory_table["$($msifileinfo.Directory)"][0])
                # Get the Default Diectory from the directory table
                $msifileinfo | Add-Member -MemberType NoteProperty -Name "DefaultDir" ($dd = $directory_table["$($msifileinfo.Directory)"][1])

                # Get the resolved path information

                #$rp = Get-Content -Path $tmprpaths | Select-String "$($msifileinfo.Component)" | Select-Object -First 1 
                $rp = $rpf | Select-String "$($msifileinfo.Component)"    | Select-Object -First 1 


                #$rp = Get-Content -Path $tmprpaths  -filter "$($msifileinfo.Component)" | Select-Object -First 1 
                $rp = $rp.ToString().split(",")[0]

                    # because the above command sometimes returns multiple values, we have to check whether we have
                    # just a string or an array, if we get an array we just take the first entry to get the path information
                    if ($rp -is [array] -eq $true)
                    {
                        $msifileinfo | Add-Member -MemberType NoteProperty -Name "TargetPath" -Value ([System.IO.Path]::GetDirectoryName($rp[0]))
                    }
                    Else
                    {
                        $msifileinfo | Add-Member -MemberType NoteProperty -Name "TargetPath" -Value ([System.IO.Path]::GetDirectoryName($rp))
                    }
                $msi_fileprops += $msifileinfo
                $fc++ 

                # to prevent write-progress from slowing down the process, only display process every 1000 milisecnds
                if ($sw.Elapsed.TotalMilliseconds -ge 1000)
                     {
                        $curmemusage = "$('{0:n2}' -f ([double](Get-Process -Id $pid).WorkingSet/1MB)) MB"
                        Write-Progress -ParentId $ParentID -Activity "Retrieving Files" -Status "Processed $fc of $totalfiles files, Memory Usage: $curmemusage" -PercentComplete (($fc / $totalfiles) * 100)
                        $sw.Reset(); $sw.Start()
                    }
            } # end while file info

            # Delete the temp resolved path output file
            Remove-item -Path "$tmprpaths" -Force

            # dump the results into the text file

            Write-Progress -ParentId $ParentID -Activity "Processing results" -Status "storing results into log file" -PercentComplete 100
            $msi_fileprops | Select-Object * | export-csv -Path $msifileinfo_output -NoClobber -NoTypeInformation -Append 
            $si++
        } 
        catch
        {
            # show the error but continue
            Write-Output "Failed to get MSI file information the error was: {0}." -f $_
        }
    } # end for each msi file
} # end process

End
{
        # remove the temp ps script
        Remove-Item -Path "$tmpps" -Force
}
} # end function