#region ScriptInfo

<#

.SYNOPSIS
Gets Windows 10 Operating System details from a computer and generated a word document with the information gathered.

.DESCRIPTION
Gets Windows 10 Operating System details from a computer and generated a word document with the information gathered. This script has various modes, GatherAndReport, GatherOnly & ReportOnly
each can be used for different purposes. The report can either be basic or detailed. This script can run against a report endpoint. See examples for futher detials.

.PARAMETER outDir
This is the directory where the xml files and reports are stored. If the directory doesn't exist it will be created.

.PARAMETER mode
[OPTIONAL] This is the mode the script runs in. There are 3 modes available. GatherAndReport, GatherOnly & ReportOnly. GatherAndReport will collect details and generate a word document. GatherOnly will collect
details only. ReportOnly will construct a word document from a provided xml file with previously gathered details. The default mode is GatherAndReport.

.PARAMETER endpoint
[OPTIONAL] This is the name of the remote endpoint if the script is to gather information from a remote endpoint. If this parameter is not specified then the script will run against the endpoint
that the script is running on.

.PARAMETER xmlReport
[OPTIONAL] This parameter is only used for ReportOnly mode and is the xml file with the details previously gathered using GatherAndReport or GatherOnly mode.

.PARAMETER reportMode
[OPTIONAL] This is the report mode of the script. It can either be Basic or Detailed. Basic will report a smaller set of data whilst Detailed will provide a greater level of report.

.EXAMPLE
.\Get-Win10Info.ps1 -outDir "c:\temp"
Runs the script in GatherAndReport mode which will save a basic report to c:\temp

.EXAMPLE
.\Get-Win10Info.ps1 -outDir "c:\temp" -endpoint PC01
Runs the script in GatherAndReport mode which will save a basic report to c:\temp with the details gather from the remote endpoint PC01.

.\Get-Win10Info.ps1 -outDir "c:\temp" -mode GatherOnly -endpoint PC01
Runs the script in GatherOnly and collects details from the remote endpoint PC01.

.\Get-Win10Info.ps1 -outDir "c:\temp" -mode ReportOnly -xmlFile c:\temp\PC01_201909291709.xml -ReportType Detailed
Runs the script in ReportOnly mode and generates a detailed report using GR001_201909291709.xml.

.LINK
https://github.com/gordonrankine/get-win10info

.NOTES
License:            MIT License
Compatibility:      Windows 10
Author:             Gordon Rankine
Date:               28/04/2021
Version:            1.3
PSScriptAnalyzer:  Pass (with caveat). Run ScriptAnalyzer with PSAvoidUsingWMICmdlet. WMI over CIM as WMI is more versatile than CIM.
Change Log:         Version  Date        Author          Comments
                    1.0      29/09/2019  Gordon Rankine  Initial script
                    1.1      31/10/2020  Gordon Rankine  Added Window System Assessment Tool (win32_winsat). Added blank lines to either side of script complete message.
                    1.2      21/11/2020  Gordon Rankine  Added Computer Certificates (System.Security.Cryptography.X509Certificates.X509Store).
                    1.3      28/04/2021  Gordon Rankine  Added Power Plan (win32_powerplan).

#>

#endregion ScriptInfo

#region Bindings
[cmdletbinding()]

Param(

    [Parameter(Mandatory=$True, Position=1, HelpMessage="This is the directory for the output file.")]
    [string]$outDir,
    [Parameter(Mandatory=$False, Position=2, HelpMessage="This is the mode that the script will run in. There are 3 modes: GatherAndReport, GatherOnly and ReportOnly. GatherAndReport is the default option.")]
    [ValidateSet('GatherAndReport','GatherOnly','ReportOnly')]
    [string]$mode = 'GatherAndReport',
    [Parameter(Mandatory=$False, Position=3, HelpMessage="This is the hostname of the endpoint that the data is to be collected from. If no endpoint is selected, the script will default to the local computer.")]
    [string]$endpoint = $env:COMPUTERNAME,
    [Parameter(Mandatory=$False, Position=4, HelpMessage="This is the xml report generated from either the GatherAndReport or GatherOnly modes.")]
    [string]$xmlReport,
    [Parameter(Mandatory=$False, Position=5, HelpMessage="This is the type of report to run. There are 2 reports: Basic or Detailed. Basic is the default report type.")]
    [ValidateSet('Basic','Detailed')]
    [string]$reportType = 'Basic'
)
#endregion Bindings

#region Functions
### START FUNCTIONS ###

    ### FUNCTION - CREATE DIRECTORY ###
    function fnCreateDir {

    <#

    .SYNOPSIS
    Creates a directory.

    .DESCRIPTION
        Creates a directory.

    .PARAMETER outDir
    This is the directory to be created.

    .EXAMPLE
    .\Create-Directory.ps1 -outDir "c:\test"
    Creates a directory called "test" in c:\

    .EXAMPLE
    .\Create-Directory.ps1 -outDir "\\COMP01\c$\test"
    Creates a directory called "test" in c:\ on COMP01

    .LINK
    https://github.com/gordonrankine/powershell

    .NOTES
        License:            MIT License
        Compatibility:      Windows 7 or Server 2008 and higher
        Author:             Gordon Rankine
        Date:               13/01/2019
        Version:            1.1
        PSSscriptAnalyzer:  Pass

    #>

        [CmdletBinding()]

            Param(

            # The directory to be created.
            [Parameter(Mandatory=$True, Position=0, HelpMessage='This is the directory to be created. E.g. C:\Temp')]
            [string]$outDir

            )

            # Create out directory if it doesnt exist
            if(!(Test-Path -path $outDir)){
                if(($outDir -notlike "*:\*") -and ($outDir -notlike "*\\*")){
                Write-Output "[ERROR]: $outDir is not a valid path. Script terminated."
                break
                }
                    try{
                    New-Item $outDir -type directory -Force -ErrorAction Stop | Out-Null
                    Write-Output "[INFO] Created output directory $outDir"
                    }
                    catch{
                    Write-Output "[ERROR]: There was an issue creating $outDir. Script terminated."
                    Write-Output ($_.Exception.Message)
                    Write-Output ""
                    break
                    }
            }
            # Directory already exists
            else{
            Write-Output "[INFO] $outDir already exists."
            }

    } # end fnCreateDir

    ### FUNCTION - CHECK POWERSHELL IS RUNNING AS ADMINISTRATOR ###
    function fnCheckPSAdmin {

    <#

    .SYNOPSIS
    Checks PowerShell is running as Administrator.

    .DESCRIPTION
    Checks PowerShell is running as Administrator.

    .LINK
    https://github.com/gordonrankine/powershell

    .NOTES
        License:            MIT License
        Compatibility:      Windows 7 or Server 2008 and higher
        Author:             Gordon Rankine
        Date:               19/09/2019
        Version:            1.0
        PSSscriptAnalyzer:  Pass

    #>

        try{
        $wIdCurrent = [System.Security.Principal.WindowsIdentity]::GetCurrent()
        $wPrinCurrent = New-Object System.Security.Principal.WindowsPrincipal($wIdCurrent)
        $wBdminRole = [System.Security.Principal.WindowsBuiltInRole]::Administrator

            if(!$wPrinCurrent.IsInRole($wBdminRole)){
            Write-Output "[ERROR] PowerShell is not running as administrator. Script terminated."
            Break
            }

        }

        catch{
        Write-Output "[ERROR] There was an unexpected error checking if PowerShell is running as administrator. Script terminated."
        Break
        }

    } # end fnCheckPSAdmin

#endregion Functions ### END FUNCTIONS ###

#region Initialize
Clear-Host

# Start stopwatch
$sw = [system.diagnostics.stopwatch]::StartNew()

fnCheckPSAdmin
fnCreateDir $outDir

$date = Get-Date -UFormat %Y%m%d%H%M    #%d-%m-%Y-%H%M
$outFile = "$outDir\$endpoint`_$date.xml"
$warnings = 0
$tab = "`t"

    # Validate correct parameters used for ReportOnly mode
    if(($mode -eq "ReportOnly") -and ($xmlReport -eq "")){
    Write-Output "[ERROR] ReportOnly mode needs the xml file. Please use -xmlReport parameter when running this mode."
    break
    }
    # Validate correct parameters used for GatherAndReport mode
    if(($mode -eq "GatherAndReport") -and ($xmlReport -ne "")){
    Write-Output "[ERROR] GatherAndReport mode should not have the xml file specified. If you have not specified -mode GatherAndReport explicitly, remember this is the default mode. Please remove -xmlReport parameter when running this mode."
    break
    }

Write-Output "[INFO] Script running in $mode mode."

    if(($mode -eq "GatherAndReport") -or ($mode -eq "GatherOnly")){
    $reportFile = $outFile
    }
    else{
    $reportFile = $xmlReport
    }

    if($mode -ne "ReportOnly"){

        # Check that content can be written to outfile
        try{
        Write-Output "[INFO] XML capture file is $outDir\$endpoint`_$date.xml."
        Add-Content $outfile "<?xml version=`"1.0`" encoding=`"UTF-8`"?>" -ErrorAction SilentlyContinue
        }
        catch{
        Write-Output "[ERROR] There was an unexpected error while writing the xml file. Script terminated."
        Write-Output "[ERROR] $($_.Exception.Message)."
        break
        }

#endregion Initialize

#region Gather

    #region GatherGeneral
    Add-Content $outfile "<info>"
    Add-Content $outfile "$tab<general>"
    Add-Content $outfile "$tab$tab<ge_server>$endpoint</ge_server>"
    Add-Content $outfile "$tab$tab<ge_inventdate>$date</ge_inventdate>"
    Add-Content $outfile "$tab$tab<ge_scriptname>$($myInvocation.MyCommand.Name)</ge_scriptname>"
    Add-Content $outfile "$tab$tab<ge_runby>$($env:USERDOMAIN + "\" + $env:USERNAME)</ge_runby>"

        if($env:COMPUTERNAME -eq $endpoint){
        $remote = "No"
        }
        else{
        $remote = "Yes"
        }

    Add-Content $outfile "$tab$tab<ge_remote>$($remote)</ge_remote>"
    Add-Content $outfile "$tab</general>"
    #endregion GatherGeneral

        #region GatherRegistryOperatingSystem
        try{
        Write-Output "[INFO] Getting Operating System details from registry."
        $reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $endpoint)
        $key = "SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion"
        $openSubKey = $reg.OpenSubKey($key)
        Add-Content $outfile "$tab<winver>"
        Add-Content $outfile "$tab$tab<version>$($openSubKey.getvalue("ReleaseId"))</version>"
        Add-Content $outfile "$tab$tab<build>$($openSubKey.getvalue("CurrentBuild") + "." + $openSubKey.getvalue("UBR"))</build>"
        Add-Content $outfile "$tab$tab<buildbranch>$($openSubKey.getvalue("BuildBranch"))</buildbranch>"
        Add-Content $outfile "$tab$tab<editionid>$($openSubKey.getvalue("EditionId"))</editionid>"
        Add-Content $outfile "$tab$tab<productname>$($openSubKey.getvalue("ProductName"))</productname>"
        Add-Content $outfile "$tab$tab<registeredorganization>$($openSubKey.getvalue("RegisteredOrganization"))</registeredorganization>"
        Add-Content $outfile "$tab$tab<registeredowner>$($openSubKey.getvalue("RegisteredOwner"))</registeredowner>"
        Add-Content $outfile "$tab</winver>"
        }
        catch{
        Add-Content $outfile "$tab<winver>"
        Add-Content $outfile "$tab$tab<errorcode>1</errorcode>"
        Add-Content $outfile "$tab$tab<errortext>$_.Exception.Message</errortext>"
        Add-Content $outfile "$tab</winver>"
        Write-Output "[WARNING] There was an unexpected error while getting the winver details. Moving on to next section..."
        $warnings ++
        }
        #endregion GatherRegistryOperatingSystem

        #region GatherWMIOperatingSystem
        $class = "win32_operatingsystem" # Single item class
        try{
        Write-Output "[INFO] Getting details from class $class."
        $props = ('buildnumber', 'caption', 'csname', 'encryptionlevel', 'installdate', 'operatingsystemsku', 'osarchitecture', 'osproductsuite', 'producttype', 'servicepackmajorversion', 'servicepackminorversion', 'totalvisiblememorysize', 'version')
        $wmi = Get-WmiObject -Namespace root/cimv2 -Class $class -ComputerName $endpoint -ErrorAction Stop | Select-Object $props
        Add-Content $outfile "$tab<$class>"
            foreach($prop in $props){
            Add-Content $outfile "$tab$tab<$prop>$($wmi.$prop)</$prop>"
            }
        Add-Content $outfile "$tab</$class>"
        }
        catch{
        Add-Content $outfile "$tab<$class>"
        Add-Content $outfile "$tab$tab<errorcode>1</errorcode>"
        Add-Content $outfile "$tab$tab<errortext>$_.Exception.Message</errortext>"
        Add-Content $outfile "$tab</$class>"
        Write-Output "[WARNING] There was an unexpected error while getting the $class class. Moving on to next section..."
        $warnings ++
        }
        #endregion GatherWMIOperatingSystem

        #region GatherWMIComputerSystem
        $class = "win32_computersystem" # Single item class
        try{
        Write-Output "[INFO] Getting details from class $class."
        $props = ('adminpasswordstatus', 'name', 'domain', 'domainrole', 'manufacturer', 'model', 'numberoflogicalprocessors', 'partofdomain', 'roles', 'systemtype', 'totalphysicalmemory', 'wakeuptype')
        $wmi = Get-WmiObject -Namespace root/cimv2 -Class $class -ComputerName $endpoint -ErrorAction Stop | Select-Object $props
        Add-Content $outfile "$tab<$class>"
            foreach($prop in $props){
            Add-Content $outfile "$tab$tab<$prop>$($wmi.$prop)</$prop>"
            }
        Add-Content $outfile "$tab</$class>"
        }
        catch{
        Add-Content $outfile "$tab<$class>"
        Add-Content $outfile "$tab$tab<errorcode>1</errorcode>"
        Add-Content $outfile "$tab$tab<errortext>$_.Exception.Message</errortext>"
        Add-Content $outfile "$tab</$class>"
        Write-Output "[WARNING] There was an unexpected error while getting the $class class. Moving on to next section..."
        $warnings ++
        }
        #endregion GatherWMIComputerSystem

        #region GatherWMIWinSat
        $class = "win32_winsat" # Single item class
        try{
        Write-Output "[INFO] Getting details from class $class."
        $props = ('cpuscore', 'd3dscore', 'diskscore', 'graphicsscore', 'memoryscore', 'timetaken', 'winsatassessmentstate', 'winsprlevel')
        $wmi = Get-WmiObject -Namespace root/cimv2 -Class $class -ComputerName $endpoint -ErrorAction Stop | Select-Object $props
        Add-Content $outfile "$tab<$class>"
            foreach($prop in $props){
            Add-Content $outfile "$tab$tab<$prop>$($wmi.$prop)</$prop>"
            }
        Add-Content $outfile "$tab</$class>"
        }
        catch{
        Add-Content $outfile "$tab<$class>"
        Add-Content $outfile "$tab$tab<errorcode>1</errorcode>"
        Add-Content $outfile "$tab$tab<errortext>$_.Exception.Message</errortext>"
        Add-Content $outfile "$tab</$class>"
        Write-Output "[WARNING] There was an unexpected error while getting the $class class. Moving on to next section..."
        $warnings ++
        }
        #endregion GatherWMIWinSat

        #region GatherWMINetworkAdapterConfiguration
        $class = "win32_networkadapterconfiguration" # Multi item class
        try{
        Write-Output "[INFO] Getting details from class $class."
        $props = ('dhcpleaseexpires', 'description', 'dhcpenabled', 'dhcpleaseobtained', 'dhcpserver', 'dnsdomain', 'dnsdomainsuffixsearchorder' ,'dnsenabledforwinsresolution', 'dnshostname', 'dnsserversearchorder', 'ipaddress', 'ipenabled', 'ipfiltersecurityenabled', 'winsenablelmhostslookup', 'winsprimaryserver', 'winssecondaryserver', 'caption', 'defaultipgateway', 'ipsubnet', 'macaddress')
        $wmi = Get-WmiObject -Namespace root/cimv2 -Class $class -Filter "IPEnabled = 'True'" -ComputerName $endpoint -ErrorAction Stop | Select-Object $props | Sort-Object description
        Add-Content $outfile "$tab<$class>"
            foreach($wm in $wmi){
            Add-Content $outfile "$tab$tab<$class`_multi>"
                Add-Content $outfile "$tab$tab$tab<dhcpleaseexpires>$($wm.dhcpleaseexpires)</dhcpleaseexpires>"
                Add-Content $outfile "$tab$tab$tab<description>$($wm.description)</description>"
                Add-Content $outfile "$tab$tab$tab<dhcpenabled>$($wm.dhcpenabled)</dhcpenabled>"
                Add-Content $outfile "$tab$tab$tab<dhcpleaseobtained>$($wm.dhcpleaseobtained)</dhcpleaseobtained>"
                Add-Content $outfile "$tab$tab$tab<dhcpserver>$($wm.dhcpserver)</dhcpserver>"
                Add-Content $outfile "$tab$tab$tab<dnsdomain>$($wm.dnsdomain)</dnsdomain>"
                Add-Content $outfile "$tab$tab$tab<dnsdomainsuffixsearchorder>$($wm.dnsdomainsuffixsearchorder)</dnsdomainsuffixsearchorder>"
                Add-Content $outfile "$tab$tab$tab<dnsenabledforwinsresolution>$($wm.dnsenabledforwinsresolution)</dnsenabledforwinsresolution>"
                Add-Content $outfile "$tab$tab$tab<dnshostname>$($wm.dnshostname)</dnshostname>"
                Add-Content $outfile "$tab$tab$tab<dnsserversearchorder>$($wm.dnsserversearchorder)</dnsserversearchorder>"
                Add-Content $outfile "$tab$tab$tab<ipaddress>$($wm.ipaddress)</ipaddress>"
                Add-Content $outfile "$tab$tab$tab<ipenabled>$($wm.ipenabled)</ipenabled>"
                Add-Content $outfile "$tab$tab$tab<ipfiltersecurityenabled>$($wm.ipfiltersecurityenabled)</ipfiltersecurityenabled>"
                Add-Content $outfile "$tab$tab$tab<winsenablelmhostslookup>$($wm.winsenablelmhostslookup)</winsenablelmhostslookup>"
                Add-Content $outfile "$tab$tab$tab<winsprimaryserver>$($wm.winsprimaryserver)</winsprimaryserver>"
                Add-Content $outfile "$tab$tab$tab<winssecondaryserver>$($wm.winssecondaryserver)</winssecondaryserver>"
                Add-Content $outfile "$tab$tab$tab<caption>$($wm.caption)</caption>"
                Add-Content $outfile "$tab$tab$tab<defaultipgateway>$($wm.defaultipgateway)</defaultipgateway>"
                Add-Content $outfile "$tab$tab$tab<ipsubnet>$($wm.ipsubnet)</ipsubnet>"
                Add-Content $outfile "$tab$tab$tab<macaddress>$($wm.macaddress)</macaddress>"
            Add-Content $outfile "$tab$tab</$class`_multi>"
            }
        Add-Content $outfile "$tab</$class>"
        }
        catch{
        Add-Content $outfile "$tab<$class>"
        Add-Content $outfile "$tab$tab<errorcode>1</errorcode>"
        Add-Content $outfile "$tab$tab<errortext>$_.Exception.Message</errortext>"
        Add-Content $outfile "$tab</$class>"
        Write-Output "[WARNING] There was an unexpected error while getting the $class class. Moving on to next section..."
        $warnings ++
        }
        #endregion GatherWMINetworkAdapterConfiguration

        #region GatherWMIPhysicalMemory
        $class = "win32_physicalmemory" # Multi item class
        try{
        Write-Output "[INFO] Getting details from class $class."
        $props = ('capacity', 'configuredclockspeed', 'datawidth', 'devicelocator', 'formfactor', 'manufacturer', 'memorytype', 'model', 'partnumber', 'serialnumber', 'speed')
        $wmi = Get-WmiObject -Namespace root/cimv2 -Class $class -ComputerName $endpoint -ErrorAction Stop | Select-Object $props | Sort-Object devicelocator
        Add-Content $outfile "$tab<$class>"
            foreach($wm in $wmi){
            Add-Content $outfile "$tab$tab<$class`_multi>"
                Add-Content $outfile "$tab$tab$tab<capacity>$($wm.capacity)</capacity>"
                Add-Content $outfile "$tab$tab$tab<configuredclockspeed>$($wm.configuredclockspeed)</configuredclockspeed>"
                Add-Content $outfile "$tab$tab$tab<datawidth>$($wm.datawidth)</datawidth>"
                Add-Content $outfile "$tab$tab$tab<devicelocator>$($wm.devicelocator)</devicelocator>"
                Add-Content $outfile "$tab$tab$tab<formfactor>$($wm.formfactor)</formfactor>"
                Add-Content $outfile "$tab$tab$tab<manufacturer>$($wm.manufacturer)</manufacturer>"
                Add-Content $outfile "$tab$tab$tab<memorytype>$($wm.memorytype)</memorytype>"
                Add-Content $outfile "$tab$tab$tab<model>$($wm.model)</model>"
                Add-Content $outfile "$tab$tab$tab<partnumber>$($wm.partnumber)</partnumber>"
                Add-Content $outfile "$tab$tab$tab<serialnumber>$($wm.serialnumber)</serialnumber>"
                Add-Content $outfile "$tab$tab$tab<speed>$($wm.speed)</speed>"
            Add-Content $outfile "$tab$tab</$class`_multi>"
            }
        Add-Content $outfile "$tab</$class>"
        }
        catch{
        Add-Content $outfile "$tab<$class>"
        Add-Content $outfile "$tab$tab<errorcode>1</errorcode>"
        Add-Content $outfile "$tab$tab<errortext>$_.Exception.Message</errortext>"
        Add-Content $outfile "$tab</$class>"
        Write-Output "[WARNING] There was an unexpected error while getting the $class class. Moving on to next section..."
        $warnings ++
        }
        #endregion GatherWMIPhysicalMemory

        #region GatherWMIProcessor
        $class = "win32_processor" # Multi item class
        try{
        Write-Output "[INFO] Getting details from class $class."
        $props = ('addresswidth', 'caption', 'cpustatus', 'currentclockspeed', 'datawidth', 'deviceid', 'extclock', 'l2cachesize',  'l2cachespeed', 'l3cachesize', 'l3cachespeed', 'loadpercentage', 'manufacturer', 'maxclockspeed', 'name', 'numberofcores', 'numberofenabledcore', 'numberoflogicalprocessors', 'partnumber', 'powermanagementsupported', 'processortype', 'serialnumber', 'threadcount', 'vmmonitormodeextensions', 'virtualizationfirmwareenabled')
        $wmi = Get-WmiObject -Namespace root/cimv2 -Class $class -ComputerName $endpoint -ErrorAction Stop | Select-Object $props | Sort-Object deviceid
        Add-Content $outfile "$tab<$class>"
            foreach($wm in $wmi){
            Add-Content $outfile "$tab$tab<$class`_multi>"
                Add-Content $outfile "$tab$tab$tab<addresswidth>$($wm.addresswidth)</addresswidth>"
                Add-Content $outfile "$tab$tab$tab<caption>$($wm.caption)</caption>"
                Add-Content $outfile "$tab$tab$tab<cpustatus>$($wm.cpustatus)</cpustatus>"
                Add-Content $outfile "$tab$tab$tab<currentclockspeed>$($wm.currentclockspeed)</currentclockspeed>"
                Add-Content $outfile "$tab$tab$tab<datawidth>$($wm.datawidth)</datawidth>"
                Add-Content $outfile "$tab$tab$tab<deviceid>$($wm.deviceid)</deviceid>"
                Add-Content $outfile "$tab$tab$tab<extclock>$($wm.extclock)</extclock>"
                Add-Content $outfile "$tab$tab$tab<l2cachesize>$($wm.l2cachesize)</l2cachesize>"
                Add-Content $outfile "$tab$tab$tab<l2cachespeed>$($wm.l2cachespeed)</l2cachespeed>"
                Add-Content $outfile "$tab$tab$tab<l3cachesize>$($wm.l3cachesize)</l3cachesize>"
                Add-Content $outfile "$tab$tab$tab<l3cachespeed>$($wm.l3cachespeed)</l3cachespeed>"
                Add-Content $outfile "$tab$tab$tab<loadpercentage>$($wm.loadpercentage)</loadpercentage>"
                Add-Content $outfile "$tab$tab$tab<manufacturer>$($wm.manufacturer)</manufacturer>"
                Add-Content $outfile "$tab$tab$tab<maxclockspeed>$($wm.maxclockspeed)</maxclockspeed>"
                Add-Content $outfile "$tab$tab$tab<name>$($wm.name)</name>"
                Add-Content $outfile "$tab$tab$tab<numberofcores>$($wm.numberofcores)</numberofcores>"
                Add-Content $outfile "$tab$tab$tab<numberofenabledcore>$($wm.numberofenabledcore)</numberofenabledcore>"
                Add-Content $outfile "$tab$tab$tab<numberoflogicalprocessors>$($wm.numberoflogicalprocessors)</numberoflogicalprocessors>"
                Add-Content $outfile "$tab$tab$tab<partnumber>$($wm.partnumber)</partnumber>"
                Add-Content $outfile "$tab$tab$tab<powermanagementsupported>$($wm.powermanagementsupported)</powermanagementsupported>"
                Add-Content $outfile "$tab$tab$tab<processortype>$($wm.processortype)</processortype>"
                Add-Content $outfile "$tab$tab$tab<serialnumber>$($wm.serialnumber)</serialnumber>"
                Add-Content $outfile "$tab$tab$tab<threadcount>$($wm.threadcount)</threadcount>"
                Add-Content $outfile "$tab$tab$tab<vmmonitormodeextensions>$($wm.vmmonitormodeextensions)</vmmonitormodeextensions>"
                Add-Content $outfile "$tab$tab$tab<virtualizationfirmwareenabled>$($wm.virtualizationfirmwareenabled)</virtualizationfirmwareenabled>"
            Add-Content $outfile "$tab$tab</$class`_multi>"
            }
        Add-Content $outfile "$tab</$class>"
        }
        catch{
        Add-Content $outfile "$tab<$class>"
        Add-Content $outfile "$tab$tab<errorcode>1</errorcode>"
        Add-Content $outfile "$tab$tab<errortext>$_.Exception.Message</errortext>"
        Add-Content $outfile "$tab</$class>"
        Write-Output "[WARNING] There was an unexpected error while getting the $class class. Moving on to next section..."
        $warnings ++
        }
        #endregion GatherWMIProcessor

        #region GatherWMILogicalDisk
        $class = "win32_logicaldisk" # Multi item class
        try{
        Write-Output "[INFO] Getting details from class $class."
        $props = ('deviceid', 'description', 'drivetype', 'filesystem', 'freespace', 'mediatype', 'size', 'volumename', 'volumeserialnumber')
        $wmi = Get-WmiObject -Namespace root/cimv2 -Class $class -ComputerName $endpoint -ErrorAction Stop | Select-Object $props | Sort-Object deviceid
        Add-Content $outfile "$tab<$class>"
            foreach($wm in $wmi){
            Add-Content $outfile "$tab$tab<$class`_multi>"
                Add-Content $outfile "$tab$tab$tab<deviceid>$($wm.deviceid)</deviceid>"
                Add-Content $outfile "$tab$tab$tab<description>$($wm.description)</description>"
                Add-Content $outfile "$tab$tab$tab<drivetype>$($wm.drivetype)</drivetype>"
                Add-Content $outfile "$tab$tab$tab<filesystem>$($wm.filesystem)</filesystem>"
                Add-Content $outfile "$tab$tab$tab<freespace>$($wm.freespace)</freespace>"
                Add-Content $outfile "$tab$tab$tab<mediatype>$($wm.mediatype)</mediatype>"
                Add-Content $outfile "$tab$tab$tab<size>$($wm.size)</size>"
                Add-Content $outfile "$tab$tab$tab<volumename>$($wm.volumename)</volumename>"
                Add-Content $outfile "$tab$tab$tab<volumeserialnumber>$($wm.volumeserialnumber)</volumeserialnumber>"
            Add-Content $outfile "$tab$tab</$class`_multi>"
            }
        Add-Content $outfile "$tab</$class>"
        }
        catch{
        Add-Content $outfile "$tab<$class>"
        Add-Content $outfile "$tab$tab<errorcode>1</errorcode>"
        Add-Content $outfile "$tab$tab<errortext>$_.Exception.Message</errortext>"
        Add-Content $outfile "$tab</$class>"
        Write-Output "[WARNING] There was an unexpected error while getting the $class class. Moving on to next section..."
        $warnings ++
        }
        #endregion GatherWMILogicalDisk

        #region GatherWMIDiskPartition
        $class = "win32_diskpartition" # Multi item class
        try{
        Write-Output "[INFO] Getting details from class $class."
        $props = ('blocksize', 'bootable', 'bootpartition', 'description', 'deviceid', 'diskindex', 'numberofblocks', 'primarypartition', 'size')
        $wmi = Get-WmiObject -Namespace root/cimv2 -Class $class -ComputerName $endpoint -ErrorAction Stop | Select-Object $props | Sort-Object diskindex
        Add-Content $outfile "$tab<$class>"
            foreach($wm in $wmi){
            Add-Content $outfile "$tab$tab<$class`_multi>"
                Add-Content $outfile "$tab$tab$tab<blocksize>$($wm.blocksize)</blocksize>"
                Add-Content $outfile "$tab$tab$tab<bootable>$($wm.bootable)</bootable>"
                Add-Content $outfile "$tab$tab$tab<bootpartition>$($wm.bootpartition)</bootpartition>"
                Add-Content $outfile "$tab$tab$tab<description>$($wm.description)</description>"
                Add-Content $outfile "$tab$tab$tab<deviceid>$($wm.deviceid)</deviceid>"
                Add-Content $outfile "$tab$tab$tab<diskindex>$($wm.diskindex)</diskindex>"
                Add-Content $outfile "$tab$tab$tab<numberofblocks>$($wm.numberofblocks)</numberofblocks>"
                Add-Content $outfile "$tab$tab$tab<primarypartition>$($wm.primarypartition)</primarypartition>"
                Add-Content $outfile "$tab$tab$tab<size>$($wm.size)</size>"
            Add-Content $outfile "$tab$tab</$class`_multi>"
            }
        Add-Content $outfile "$tab</$class>"
        }
        catch{
        Add-Content $outfile "$tab<$class>"
        Add-Content $outfile "$tab$tab<errorcode>1</errorcode>"
        Add-Content $outfile "$tab$tab<errortext>$_.Exception.Message</errortext>"
        Add-Content $outfile "$tab</$class>"
        Write-Output "[WARNING] There was an unexpected error while getting the $class class. Moving on to next section..."
        $warnings ++
        }
        #endregion GatherWMIDiskPartition

        #region GatherWMIShare
        $class = "win32_share" # Multi item class
        try{
        Write-Output "[INFO] Getting details from class $class."
        $props = ('caption', 'description', 'name', 'path', 'type')
        $wmi = Get-WmiObject -Namespace root/cimv2 -Class $class -ComputerName $endpoint -ErrorAction Stop | Select-Object $props | Sort-Object name
        Add-Content $outfile "$tab<$class>"
            foreach($wm in $wmi){
            Add-Content $outfile "$tab$tab<$class`_multi>"
                Add-Content $outfile "$tab$tab$tab<caption>$($wm.caption)</caption>"
                Add-Content $outfile "$tab$tab$tab<description>$($wm.description)</description>"
                Add-Content $outfile "$tab$tab$tab<name>$($wm.name)</name>"
                Add-Content $outfile "$tab$tab$tab<path>$($wm.path)</path>"
                Add-Content $outfile "$tab$tab$tab<type>$($wm.type)</type>"
            Add-Content $outfile "$tab$tab</$class`_multi>"
            }
        Add-Content $outfile "$tab</$class>"
        }
        catch{
        Add-Content $outfile "$tab<$class>"
        Add-Content $outfile "$tab$tab<errorcode>1</errorcode>"
        Add-Content $outfile "$tab$tab<errortext>$_.Exception.Message</errortext>"
        Add-Content $outfile "$tab</$class>"
        Write-Output "[WARNING] There was an unexpected error while getting the $class class. Moving on to next section..."
        $warnings ++
        }
        #endregion GatherWMIShare

        #region GatherWMIStartUpCommand
        $class = "win32_startupcommand" # Multi item class
        try{
        Write-Output "[INFO] Getting details from class $class."
        $props = ('caption', 'command', 'description', 'name', 'user')
        $wmi = Get-WmiObject -Namespace root/cimv2 -Class $class -ComputerName $endpoint -ErrorAction Stop | Select-Object $props | Sort-Object name
        Add-Content $outfile "$tab<$class>"
            foreach($wm in $wmi){
            Add-Content $outfile "$tab$tab<$class`_multi>"
                Add-Content $outfile "$tab$tab$tab<caption>$($wm.caption)</caption>"
                Add-Content $outfile "$tab$tab$tab<command>$($wm.command)</command>"
                Add-Content $outfile "$tab$tab$tab<description>$($wm.description)</description>"
                Add-Content $outfile "$tab$tab$tab<name>$($wm.name)</name>"
                Add-Content $outfile "$tab$tab$tab<user>$($wm.user)</user>"
            Add-Content $outfile "$tab$tab</$class`_multi>"
            }
        Add-Content $outfile "$tab</$class>"
        }
        catch{
        Add-Content $outfile "$tab<$class>"
        Add-Content $outfile "$tab$tab<errorcode>1</errorcode>"
        Add-Content $outfile "$tab$tab<errortext>$_.Exception.Message</errortext>"
        Add-Content $outfile "$tab</$class>"
        Write-Output "[WARNING] There was an unexpected error while getting the $class class. Moving on to next section..."
        $warnings ++
        }
        #endregion GatherWMIStartUpCommand

        #region GatherWMIPageFileUsage
        $class = "win32_pagefileusage" # Multi item class
        try{
        Write-Output "[INFO] Getting details from class $class."
        $props = ('currentusage', 'allocatedbasesize', 'caption', 'description', 'name')
        $wmi = Get-WmiObject -Namespace root/cimv2 -Class $class -ComputerName $endpoint -ErrorAction Stop | Select-Object $props | Sort-Object name
        Add-Content $outfile "$tab<$class>"
            foreach($wm in $wmi){
            Add-Content $outfile "$tab$tab<$class`_multi>"
                Add-Content $outfile "$tab$tab$tab<currentusage>$($wm.currentusage)</currentusage>"
                Add-Content $outfile "$tab$tab$tab<allocatedbasesize>$($wm.allocatedbasesize)</allocatedbasesize>"
                Add-Content $outfile "$tab$tab$tab<caption>$($wm.caption)</caption>"
                Add-Content $outfile "$tab$tab$tab<description>$($wm.description)</description>"
                Add-Content $outfile "$tab$tab$tab<name>$($wm.name)</name>"
            Add-Content $outfile "$tab$tab</$class`_multi>"
            }
        Add-Content $outfile "$tab</$class>"
        }
        catch{
        Add-Content $outfile "$tab<$class>"
        Add-Content $outfile "$tab$tab<errorcode>1</errorcode>"
        Add-Content $outfile "$tab$tab<errortext>$_.Exception.Message</errortext>"
        Add-Content $outfile "$tab</$class>"
        Write-Output "[WARNING] There was an unexpected error while getting the $class class. Moving on to next section..."
        $warnings ++
        }
        #endregion GatherWMIPageFileUsage

        #region GatherWMIQuickFixEngineering
        $class = "win32_quickfixengineering" # Multi item class
        try{
        Write-Output "[INFO] Getting details from class $class."
        $props = ('caption', 'description', 'hotfixid', 'installedby', 'installedon')
        $wmi = Get-WmiObject -Namespace root/cimv2 -Class $class -ComputerName $endpoint -ErrorAction Stop | Select-Object $props | Sort-Object hotfixid
        Add-Content $outfile "$tab<$class>"
            foreach($wm in $wmi){
            Add-Content $outfile "$tab$tab<$class`_multi>"
                Add-Content $outfile "$tab$tab$tab<caption>$($wm.caption)</caption>"
                Add-Content $outfile "$tab$tab$tab<description>$($wm.description)</description>"
                Add-Content $outfile "$tab$tab$tab<hotfixid>$($wm.hotfixid)</hotfixid>"
                Add-Content $outfile "$tab$tab$tab<installedby>$($wm.installedby)</installedby>"
                Add-Content $outfile "$tab$tab$tab<installedon>$($wm.installedon)</installedon>"
            Add-Content $outfile "$tab$tab</$class`_multi>"
            }
        Add-Content $outfile "$tab</$class>"
        }
        catch{
        Add-Content $outfile "$tab<$class>"
        Add-Content $outfile "$tab$tab<errorcode>1</errorcode>"
        Add-Content $outfile "$tab$tab<errortext>$_.Exception.Message</errortext>"
        Add-Content $outfile "$tab</$class>"
        Write-Output "[WARNING] There was an unexpected error while getting the $class class. Moving on to next section..."
        $warnings ++
        }
        #endregion GatherWMIQuickFixEngineering

        #region GatherWMISystemEnclosure
        $class = "win32_systemenclosure" # Single item class
        try{
        Write-Output "[INFO] Getting details from class $class."
        $props = ('chassistypes', 'lockpresent', 'manufacturer', 'model', 'securitystatus', 'serialnumber')
        $wmi = Get-WmiObject -Namespace root/cimv2 -Class $class -ComputerName $endpoint -ErrorAction Stop | Select-Object $props
        Add-Content $outfile "$tab<$class>"
            foreach($prop in $props){
            Add-Content $outfile "$tab$tab<$prop>$($wmi.$prop)</$prop>"
            }
        Add-Content $outfile "$tab</$class>"
        }
        catch{
        Add-Content $outfile "$tab<$class>"
        Add-Content $outfile "$tab$tab<errorcode>1</errorcode>"
        Add-Content $outfile "$tab$tab<errortext>$_.Exception.Message</errortext>"
        Add-Content $outfile "$tab</$class>"
        Write-Output "[WARNING] There was an unexpected error while getting the $class class. Moving on to next section..."
        $warnings ++
        }
        #endregion GatherWMISystemEnclosure

        #region GatherWMIBootConfiguration
        $class = "win32_bootconfiguration" # Single item class
        try{
        Write-Output "[INFO] Getting details from class $class."
        $props = ('bootdirectory', 'caption', 'configurationpath', 'description', 'scratchdirectory')
        $wmi = Get-WmiObject -Namespace root/cimv2 -Class $class -ComputerName $endpoint -ErrorAction Stop | Select-Object $props
        Add-Content $outfile "$tab<$class>"
            foreach($prop in $props){
            Add-Content $outfile "$tab$tab<$prop>$($wmi.$prop)</$prop>"
            }
        Add-Content $outfile "$tab</$class>"
        }
        catch{
        Add-Content $outfile "$tab<$class>"
        Add-Content $outfile "$tab$tab<errorcode>1</errorcode>"
        Add-Content $outfile "$tab$tab<errortext>$_.Exception.Message</errortext>"
        Add-Content $outfile "$tab</$class>"
        Write-Output "[WARNING] There was an unexpected error while getting the $class class. Moving on to next section..."
        $warnings ++
        }
        #endregion GatherWMIBootConfiguration

        #region GatherWMIBios
        $class = "win32_bios" # Single item class
        try{
        Write-Output "[INFO] Getting details from class $class."
        $props = ('biosversion', 'bioscharacteristics', 'buildnumber', 'caption', 'description', 'manufacturer', 'name', 'primarybios', 'releasedate', 'smbiosbiosversion', 'smbiosmajorversion', 'smbiosminorversion', 'smbiospresent', 'serialnumber', 'softwareelementid', 'softwareelementstate', 'systembiosmajorversion', 'systembiosminorversion', 'version')
        $wmi = Get-WmiObject -Namespace root/cimv2 -Class $class -ComputerName $endpoint -ErrorAction Stop | Select-Object $props
        Add-Content $outfile "$tab<$class>"
            foreach($prop in $props){
            Add-Content $outfile "$tab$tab<$prop>$($wmi.$prop)</$prop>"
            }
        Add-Content $outfile "$tab</$class>"
        }
        catch{
        Add-Content $outfile "$tab<$class>"
        Add-Content $outfile "$tab$tab<errorcode>1</errorcode>"
        Add-Content $outfile "$tab$tab<errortext>$_.Exception.Message</errortext>"
        Add-Content $outfile "$tab</$class>"
        Write-Output "[WARNING] There was an unexpected error while getting the $class class. Moving on to next section..."
        $warnings ++
        }
        #endregion GatherWMIBios

        #region GatherWMIUserAccount
        $class = "win32_useraccount" # Multi item class
        try{
        Write-Output "[INFO] Getting details from class $class."
        $props = ('accounttype', 'caption', 'description', 'disabled', 'domain', 'fullname', 'localaccount', 'lockout', 'name', 'passwordchangeable', 'passwordexpires', 'passwordrequired', 'sid', 'sidtype')
        $wmi = Get-WmiObject -Namespace root/cimv2 -Class $class -Filter "Domain='$endpoint'" -ComputerName $endpoint -ErrorAction Stop | Select-Object $props | Sort-Object name
        Add-Content $outfile "$tab<$class>"
            foreach($wm in $wmi){
            Add-Content $outfile "$tab$tab<$class`_multi>"
                Add-Content $outfile "$tab$tab$tab<accounttype>$($wm.accounttype)</accounttype>"
                Add-Content $outfile "$tab$tab$tab<caption>$($wm.caption)</caption>"
                Add-Content $outfile "$tab$tab$tab<description>$($wm.description)</description>"
                Add-Content $outfile "$tab$tab$tab<disabled>$($wm.disabled)</disabled>"
                Add-Content $outfile "$tab$tab$tab<domain>$($wm.domain)</domain>"
                Add-Content $outfile "$tab$tab$tab<fullname>$($wm.fullname)</fullname>"
                Add-Content $outfile "$tab$tab$tab<localaccount>$($wm.localaccount)</localaccount>"
                Add-Content $outfile "$tab$tab$tab<lockout>$($wm.lockout)</lockout>"
                Add-Content $outfile "$tab$tab$tab<name>$($wm.name)</name>"
                Add-Content $outfile "$tab$tab$tab<passwordchangeable>$($wm.passwordchangeable)</passwordchangeable>"
                Add-Content $outfile "$tab$tab$tab<passwordexpires>$($wm.passwordexpires)</passwordexpires>"
                Add-Content $outfile "$tab$tab$tab<passwordrequired>$($wm.passwordrequired)</passwordrequired>"
                Add-Content $outfile "$tab$tab$tab<sid>$($wm.sid)</sid>"
                Add-Content $outfile "$tab$tab$tab<sidtype>$($wm.sidtype)</sidtype>"
            Add-Content $outfile "$tab$tab</$class`_multi>"
            }
        Add-Content $outfile "$tab</$class>"
        }
        catch{
        Add-Content $outfile "$tab<$class>"
        Add-Content $outfile "$tab$tab<errorcode>1</errorcode>"
        Add-Content $outfile "$tab$tab<errortext>$_.Exception.Message</errortext>"
        Add-Content $outfile "$tab</$class>"
        Write-Output "[WARNING] There was an unexpected error while getting the $class class. Moving on to next section..."
        $warnings ++
        }
        #endregion GatherWMIUserAccount

        #region GatherWMIGroup
        $class = "win32_group" # Multi item class
        try{
        Write-Output "[INFO] Getting details from class $class."
        $props = ('caption', 'description', 'domain', 'localaccount', 'name', 'sid', 'sidtype')
        $wmi = Get-WmiObject -Namespace root/cimv2 -Class $class -Filter "Domain='$endpoint'" -ComputerName $endpoint -ErrorAction Stop | Select-Object $props | Sort-Object name
        Add-Content $outfile "$tab<$class>"
            foreach($wm in $wmi){
            Add-Content $outfile "$tab$tab<$class`_multi>"
                Add-Content $outfile "$tab$tab$tab<caption>$($wm.caption)</caption>"
                Add-Content $outfile "$tab$tab$tab<description>$($wm.description)</description>"
                Add-Content $outfile "$tab$tab$tab<domain>$($wm.domain)</domain>"
                Add-Content $outfile "$tab$tab$tab<localaccount>$($wm.localaccount)</localaccount>"
                Add-Content $outfile "$tab$tab$tab<name>$($wm.name)</name>"
                Add-Content $outfile "$tab$tab$tab<sid>$($wm.sid)</sid>"
                Add-Content $outfile "$tab$tab$tab<sidtype>$($wm.sidtype)</sidtype>"
            Add-Content $outfile "$tab$tab</$class`_multi>"
            }
        Add-Content $outfile "$tab</$class>"
        }
        catch{
        Add-Content $outfile "$tab<$class>"
        Add-Content $outfile "$tab$tab<errorcode>1</errorcode>"
        Add-Content $outfile "$tab$tab<errortext>$_.Exception.Message</errortext>"
        Add-Content $outfile "$tab</$class>"
        Write-Output "[WARNING] There was an unexpected error while getting the $class class. Moving on to next section..."
        $warnings ++
        }
        #endregion GatherWMIGroup

        #region GatherWMIGroupMembership
        $class = "win32_group" # Multi item class
        try{
        Write-Output "[INFO] Getting details from class $class and querying membership."
        $props = ('caption', 'description', 'domain', 'localaccount', 'name', 'sid', 'sidtype')
        $groups = Get-WmiObject -Namespace root/cimv2 -Class $class -Filter "Domain='$endpoint'" -ComputerName $endpoint -ErrorAction Stop | Select-Object $props | Sort-Object name
        Add-Content $outfile "$tab<$class`_membership>"


            foreach($group in $groups){
            $query = "GroupComponent = `"$class.Domain='$($group.domain)',Name='$($group.Name)'`""
            $users = (Get-WmiObject -Namespace root/cimv2 -Class win32_groupuser -Filter $query).PartComponent
            Add-Content $outfile "$tab$tab<$class`_multi>"
            Add-Content $outfile "$tab$tab$tab<name>$($group.Domain +'\' + $group.Name)</name>"
            Add-Content $outfile "$tab$tab$tab$tab<members>"

                  foreach ($user in $users){
 		          $domain = $user.Substring($user.IndexOf("`"")+1)
                  $u = $domain
		          $domain = $domain.Substring(0,$domain.IndexOf("`""))
                  $u = $u.Substring($u.IndexOf("`"")+1)
                  $u = $u.Substring($u.IndexOf("`"")+1)
		          $u = $u.Substring(0,($u.Length-1))
                  Add-Content $outfile "$tab$tab$tab$tab$tab<member>$($domain +'\' + $u)</member>"
                  }

            Add-Content $outfile "$tab$tab$tab$tab</members>"
            Add-Content $outfile "$tab$tab</$class`_multi>"

            }

        Add-Content $outfile "$tab</$class`_membership>"

        }
        catch{
        Add-Content $outfile "$tab<$class>"
        Add-Content $outfile "$tab$tab<errorcode>1</errorcode>"
        Add-Content $outfile "$tab$tab<errortext>$_.Exception.Message</errortext>"
        Add-Content $outfile "$tab</$class>"
        Write-Output "[WARNING] There was an unexpected error while getting the $class class. Moving on to next section..."
        $warnings ++
        }
        #endregion GatherWMIGroupMembership

        #region GatherWMISystemAccounts
        $class = "win32_systemaccount" # Multi item class
        try{
        Write-Output "[INFO] Getting details from class $class."
        $props = ('caption', 'description', 'domain', 'localaccount', 'name', 'sid', 'sidtype')
        $wmi = Get-WmiObject -Namespace root/cimv2 -Class $class -Filter "Domain='$endpoint'" -ComputerName $endpoint -ErrorAction Stop | Select-Object $props | Sort-Object name
        Add-Content $outfile "$tab<$class>"
            foreach($wm in $wmi){
            Add-Content $outfile "$tab$tab<$class`_multi>"
                Add-Content $outfile "$tab$tab$tab<caption>$($wm.caption)</caption>"
                Add-Content $outfile "$tab$tab$tab<description>$($wm.description)</description>"
                Add-Content $outfile "$tab$tab$tab<domain>$($wm.domain)</domain>"
                Add-Content $outfile "$tab$tab$tab<localaccount>$($wm.localaccount)</localaccount>"
                Add-Content $outfile "$tab$tab$tab<name>$($wm.name)</name>"
                Add-Content $outfile "$tab$tab$tab<sid>$($wm.sid)</sid>"
                Add-Content $outfile "$tab$tab$tab<sidtype>$($wm.sidtype)</sidtype>"
            Add-Content $outfile "$tab$tab</$class`_multi>"
            }
        Add-Content $outfile "$tab</$class>"
        }
        catch{
        Add-Content $outfile "$tab<$class>"
        Add-Content $outfile "$tab$tab<errorcode>1</errorcode>"
        Add-Content $outfile "$tab$tab<errortext>$_.Exception.Message</errortext>"
        Add-Content $outfile "$tab</$class>"
        Write-Output "[WARNING] There was an unexpected error while getting the $class class. Moving on to next section..."
        $warnings ++
        }
        #endregion GatherWMISystemAccounts

        #region GatherWMIService
        $class = "win32_service" # Multi item class
        try{
        Write-Output "[INFO] Getting details from class $class."
        $props = ('caption', 'delayedautostart', 'description', 'displayname', 'name', 'pathname', 'started',  'startmode', 'startname', 'state')
        $wmi = Get-WmiObject -Namespace root/cimv2 -Class $class -ComputerName $endpoint -ErrorAction Stop | Select-Object $props | Sort-Object caption
        Add-Content $outfile "$tab<$class>"
            foreach($wm in $wmi){
            Add-Content $outfile "$tab$tab<$class`_multi>"
                Add-Content $outfile "$tab$tab$tab<caption>$($wm.caption)</caption>"
                Add-Content $outfile "$tab$tab$tab<delayedautostart>$($wm.delayedautostart)</delayedautostart>"
                Add-Content $outfile "$tab$tab$tab<description>$($wm.description)</description>"
                Add-Content $outfile "$tab$tab$tab<displayname>$($wm.displayname)</displayname>"
                Add-Content $outfile "$tab$tab$tab<name>$($wm.name)</name>"
                Add-Content $outfile "$tab$tab$tab<pathname>$($wm.pathname)</pathname>"
                Add-Content $outfile "$tab$tab$tab<started>$($wm.started)</started>"
                Add-Content $outfile "$tab$tab$tab<startmode>$($wm.startmode)</startmode>"
                Add-Content $outfile "$tab$tab$tab<startname>$($wm.startname)</startname>"
                Add-Content $outfile "$tab$tab$tab<state>$($wm.state)</state>"
                Add-Content $outfile "$tab$tab</$class`_multi>"
            }
        Add-Content $outfile "$tab</$class>"
        }
        catch{
        Add-Content $outfile "$tab<$class>"
        Add-Content $outfile "$tab$tab<errorcode>1</errorcode>"
        Add-Content $outfile "$tab$tab<errortext>$_.Exception.Message</errortext>"
        Add-Content $outfile "$tab</$class>"
        Write-Output "[WARNING] There was an unexpected error while getting the $class class. Moving on to next section..."
        $warnings ++
        }
        #endregion GatherWMIService

        #region GatherWMIProcess
        $class = "win32_process" # Multi item class
        try{
        Write-Output "[INFO] Getting details from class $class."
        $props = ('caption', 'description', 'name', 'processname')
        $wmi = Get-WmiObject -Namespace root/cimv2 -Class $class -ComputerName $endpoint -ErrorAction Stop | Select-Object $props | Sort-Object name
        Add-Content $outfile "$tab<$class>"
            foreach($wm in $wmi){
            Add-Content $outfile "$tab$tab<$class`_multi>"
               Add-Content $outfile "$tab$tab$tab<caption>$($wm.caption)</caption>"
               Add-Content $outfile "$tab$tab$tab<description>$($wm.description)</description>"
               Add-Content $outfile "$tab$tab$tab<name>$($wm.name)</name>"
               Add-Content $outfile "$tab$tab$tab<processname>$($wm.processname)</processname>"
            Add-Content $outfile "$tab$tab</$class`_multi>"
            }
        Add-Content $outfile "$tab</$class>"
        }
        catch{
        Add-Content $outfile "$tab<$class>"
        Add-Content $outfile "$tab$tab<errorcode>1</errorcode>"
        Add-Content $outfile "$tab$tab<errortext>$_.Exception.Message</errortext>"
        Add-Content $outfile "$tab</$class>"
        Write-Output "[WARNING] There was an unexpected error while getting the $class class. Moving on to next section..."
        $warnings ++
        }
        #endregion GatherWMIProcess

        #region GatherWMISystemDriver
        $class = "win32_systemdriver" # Multi item class
        try{
        Write-Output "[INFO] Getting details from class $class."
        $props = ('caption', 'description', 'displayname', 'name', 'servicetype', 'startmode', 'started', 'state')
        $wmi = Get-WmiObject -Namespace root/cimv2 -Class $class -ComputerName $endpoint -ErrorAction Stop | Select-Object $props | Sort-Object name
        Add-Content $outfile "$tab<$class>"
            foreach($wm in $wmi){
            Add-Content $outfile "$tab$tab<$class`_multi>"
                Add-Content $outfile "$tab$tab$tab<caption>$($wm.caption)</caption>"
                Add-Content $outfile "$tab$tab$tab<description>$($wm.description)</description>"
                Add-Content $outfile "$tab$tab$tab<displayname>$($wm.displayname)</displayname>"
                Add-Content $outfile "$tab$tab$tab<name>$($wm.name)</name>"
                Add-Content $outfile "$tab$tab$tab<servicetype>$($wm.servicetype)</servicetype>"
                Add-Content $outfile "$tab$tab$tab<startmode>$($wm.startmode)</startmode>"
                Add-Content $outfile "$tab$tab$tab<started>$($wm.started)</started>"
                Add-Content $outfile "$tab$tab$tab<state>$($wm.state)</state>"
            Add-Content $outfile "$tab$tab</$class`_multi>"
            }
        Add-Content $outfile "$tab</$class>"
        }
        catch{
        Add-Content $outfile "$tab<$class>"
        Add-Content $outfile "$tab$tab<errorcode>1</errorcode>"
        Add-Content $outfile "$tab$tab<errortext>$_.Exception.Message</errortext>"
        Add-Content $outfile "$tab</$class>"
        Write-Output "[WARNING] There was an unexpected error while getting the $class class. Moving on to next section..."
        $warnings ++
        }
        #endregion GatherWMISystemDriver

        #region GatherWMIPNPEntity
        $class = "win32_pnpentity" # Multi item class
        try{
        Write-Output "[INFO] Getting details from class $class."
        $props = ('caption', 'classguid', 'compatibleid', 'description', 'deviceid', 'hardwareid', 'manufacturer', 'name', 'pnpclass', 'pnpdeviceid', 'present', 'service')
        $wmi = Get-WmiObject -Namespace root/cimv2 -Class $class -ComputerName $endpoint -ErrorAction Stop | Select-Object $props | Sort-Object caption
        Add-Content $outfile "$tab<$class>"
            foreach($wm in $wmi){
            Add-Content $outfile "$tab$tab<$class`_multi>"
                Add-Content $outfile "$tab$tab$tab<caption>$($wm.caption)</caption>"
                Add-Content $outfile "$tab$tab$tab<classguid>$($wm.classguid)</classguid>"
                Add-Content $outfile "$tab$tab$tab<compatibleid>$($wm.compatibleid)</compatibleid>"
                Add-Content $outfile "$tab$tab$tab<description>$($wm.description)</description>"
                Add-Content $outfile "$tab$tab$tab<deviceid>$($wm.deviceid)</deviceid>"
                Add-Content $outfile "$tab$tab$tab<hardwareid>$($wm.hardwareid)</hardwareid>"
                Add-Content $outfile "$tab$tab$tab<manufacturer>$($wm.manufacturer)</manufacturer>"
                Add-Content $outfile "$tab$tab$tab<name>$($wm.name)</name>"
                Add-Content $outfile "$tab$tab$tab<pnpclass>$($wm.pnpclass)</pnpclass>"
                Add-Content $outfile "$tab$tab$tab<pnpdeviceid>$($wm.pnpdeviceid)</pnpdeviceid>"
                Add-Content $outfile "$tab$tab$tab<present>$($wm.present)</present>"
                Add-Content $outfile "$tab$tab$tab<service>$($wm.service)</service>"
                Add-Content $outfile "$tab$tab</$class`_multi>"
            }
        Add-Content $outfile "$tab</$class>"
        }
        catch{
        Add-Content $outfile "$tab<$class>"
        Add-Content $outfile "$tab$tab<errorcode>1</errorcode>"
        Add-Content $outfile "$tab$tab<errortext>$_.Exception.Message</errortext>"
        Add-Content $outfile "$tab</$class>"
        Write-Output "[WARNING] There was an unexpected error while getting the $class class. Moving on to next section..."
        $warnings ++
        }
        #endregion GatherWMIPNPEntity

        #region GatherWMITimezone
        $class = "win32_timezone" # Single item class
        try{
        Write-Output "[INFO] Getting details from class $class."
        $props = ('caption','standardname')
        $wmi = Get-WmiObject -Namespace root/cimv2 -Class $class -ComputerName $endpoint -ErrorAction Stop | Select-Object $props
        Add-Content $outfile "$tab<$class>"
            foreach($prop in $props){
            Add-Content $outfile "$tab$tab<$prop>$($wmi.$prop)</$prop>"
            }
        Add-Content $outfile "$tab</$class>"
        }
        catch{
        Add-Content $outfile "$tab<$class>"
        Add-Content $outfile "$tab$tab<errorcode>1</errorcode>"
        Add-Content $outfile "$tab$tab<errortext>$_.Exception.Message</errortext>"
        Add-Content $outfile "$tab</$class>"
        Write-Output "[WARNING] There was an unexpected error while getting the $class class. Moving on to next section..."
        $warnings ++
        }
        #endregion GatherWMITimezone

        #region GatherWMIRegistry
        $class = "win32_registry" # Single item class
        try{
        Write-Output "[INFO] Getting details from class $class."
        $props = ('currentsize','maximumsize')
        $wmi = Get-WmiObject -Namespace root/cimv2 -Class $class -ComputerName $endpoint -ErrorAction Stop | Select-Object $props
        Add-Content $outfile "$tab<$class>"
            foreach($prop in $props){
            Add-Content $outfile "$tab$tab<$prop>$($wmi.$prop)</$prop>"
            }
        Add-Content $outfile "$tab</$class>"
        }
        catch{
        Add-Content $outfile "$tab<$class>"
        Add-Content $outfile "$tab$tab<errorcode>1</errorcode>"
        Add-Content $outfile "$tab$tab<errortext>$_.Exception.Message</errortext>"
        Add-Content $outfile "$tab</$class>"
        Write-Output "[WARNING] There was an unexpected error while getting the $class class. Moving on to next section..."
        $warnings ++
        }
        #endregion GatherWMIRegistry

        #region GatherWMIEnvironment
        $class = "win32_environment" # Multi item class
        try{
        Write-Output "[INFO] Getting details from class $class."
        $props = ('name', 'systemvariable', 'caption', 'username', 'variablevalue')
        $wmi = Get-WmiObject -Namespace root/cimv2 -Class $class -ComputerName $endpoint -ErrorAction Stop | Select-Object $props | Sort-Object name
        Add-Content $outfile "$tab<$class>"
            foreach($wm in $wmi){
            Add-Content $outfile "$tab$tab<$class`_multi>"

                # replace , with ; for any possible values that could have
                $caption = $wm.caption -replace "<SYSTEM>", "SYSTEM" # Remove <> from System, it breaks the xml
                $username = $wm.username -replace "<SYSTEM>", "SYSTEM" # Remove <> from System, it breaks the xml

                Add-Content $outfile "$tab$tab$tab<name>$($wm.name)</name>"
                Add-Content $outfile "$tab$tab$tab<systemvariable>$($wm.systemvariable)</systemvariable>"
                Add-Content $outfile "$tab$tab$tab<caption>$caption</caption>"
                Add-Content $outfile "$tab$tab$tab<username>$username</username>"
                Add-Content $outfile "$tab$tab$tab<variablevalue>$($wm.variablevalue)</variablevalue>"
            Add-Content $outfile "$tab$tab</$class`_multi>"
            }
        Add-Content $outfile "$tab</$class>"
        }
        catch{
        Add-Content $outfile "$tab<$class>"
        Add-Content $outfile "$tab$tab<errorcode>1</errorcode>"
        Add-Content $outfile "$tab$tab<errortext>$_.Exception.Message</errortext>"
        Add-Content $outfile "$tab</$class>"
        Write-Output "[WARNING] There was an unexpected error while getting the $class class. Moving on to next section..."
        $warnings ++
        }
        #endregion GatherWMIEnvironment

        #region GatherWMICDRom
        $class = "win32_cdromdrive" # Multi item class
        try{
        Write-Output "[INFO] Getting details from class $class."
        $props = ('availability', 'capabilitydescriptions', 'description', 'deviceid', 'drive', 'manufacturer', 'mediatype', 'name', 'pnpdeviceid', 'serialnumber')
        $wmi = Get-WmiObject -Namespace root/cimv2 -Class $class -ComputerName $endpoint -ErrorAction Stop | Select-Object $props | Sort-Object name
        Add-Content $outfile "$tab<$class>"
            foreach($wm in $wmi){
            Add-Content $outfile "$tab$tab<$class`_multi>"
                Add-Content $outfile "$tab$tab$tab<availability>$($wm.availability)</availability>"
                Add-Content $outfile "$tab$tab$tab<capabilitydescriptions>$($wm.capabilitydescriptions)</capabilitydescriptions>"
                Add-Content $outfile "$tab$tab$tab<description>$($wm.description)</description>"
                Add-Content $outfile "$tab$tab$tab<deviceid>$($wm.deviceid)</deviceid>"
                Add-Content $outfile "$tab$tab$tab<drive>$($wm.drive)</drive>"
                Add-Content $outfile "$tab$tab$tab<manufacturer>$($wm.manufacturer)</manufacturer>"
                Add-Content $outfile "$tab$tab$tab<mediatype>$($wm.mediatype)</mediatype>"
                Add-Content $outfile "$tab$tab$tab<name>$($wm.name)</name>"
                Add-Content $outfile "$tab$tab$tab<pnpdeviceid>$($wm.pnpdeviceid)</pnpdeviceid>"
                Add-Content $outfile "$tab$tab$tab<serialnumber>$($wm.serialnumber)</serialnumber>"
            Add-Content $outfile "$tab$tab</$class`_multi>"
            }
        Add-Content $outfile "$tab</$class>"
        }
        catch{
        Add-Content $outfile "$tab<$class>"
        Add-Content $outfile "$tab$tab<errorcode>1</errorcode>"
        Add-Content $outfile "$tab$tab<errortext>$_.Exception.Message</errortext>"
        Add-Content $outfile "$tab</$class>"
        Write-Output "[WARNING] There was an unexpected error while getting the $class class. Moving on to next section..."
        $warnings ++
        }
        #endregion GatherWMICDRom

        #region GatherWMIVideoController
        $class = "win32_videocontroller" # Multi item class
        try{
        Write-Output "[INFO] Getting details from class $class."
        $props = ('adaptercompatibility', 'adapterdactype', 'adapterram', 'availability', 'deviceid', 'driverdate', 'driverversion', 'inffilename','infsection', 'installeddisplaydrivers', 'name', 'pnpdeviceid', 'videomodedescription', 'videoprocessor')
        $wmi = Get-WmiObject -Namespace root/cimv2 -Class $class -ComputerName $endpoint -ErrorAction Stop | Select-Object $props | Sort-Object deviceid
        Add-Content $outfile "$tab<$class>"
            foreach($wm in $wmi){
            Add-Content $outfile "$tab$tab<$class`_multi>"
                Add-Content $outfile "$tab$tab$tab<adaptercompatibility>$($wm.adaptercompatibility)</adaptercompatibility>"
                Add-Content $outfile "$tab$tab$tab<adapterdactype>$($wm.adapterdactype)</adapterdactype>"
                Add-Content $outfile "$tab$tab$tab<adapterram>$($wm.adapterram)</adapterram>"
                Add-Content $outfile "$tab$tab$tab<availability>$($wm.availability)</availability>"
                Add-Content $outfile "$tab$tab$tab<deviceid>$($wm.deviceid)</deviceid>"
                Add-Content $outfile "$tab$tab$tab<driverdate>$($wm.driverdate)</driverdate>"
                Add-Content $outfile "$tab$tab$tab<driverversion>$($wm.driverversion)</driverversion>"
                Add-Content $outfile "$tab$tab$tab<inffilename>$($wm.inffilename)</inffilename>"
                Add-Content $outfile "$tab$tab$tab<infsection>$($wm.infsection)</infsection>"
                Add-Content $outfile "$tab$tab$tab<installeddisplaydrivers>$($wm.installeddisplaydrivers)</installeddisplaydrivers>"
                Add-Content $outfile "$tab$tab$tab<name>$($wm.name)</name>"
                Add-Content $outfile "$tab$tab$tab<pnpdeviceid>$($wm.pnpdeviceid)</pnpdeviceid>"
                Add-Content $outfile "$tab$tab$tab<videomodedescription>$($wm.videomodedescription)</videomodedescription>"
                Add-Content $outfile "$tab$tab$tab<videoprocessor>$($wm.videoprocessor)</videoprocessor>"
            Add-Content $outfile "$tab$tab</$class`_multi>"
            }
        Add-Content $outfile "$tab</$class>"
        }
        catch{
        Add-Content $outfile "$tab<$class>"
        Add-Content $outfile "$tab$tab<errorcode>1</errorcode>"
        Add-Content $outfile "$tab$tab<errortext>$_.Exception.Message</errortext>"
        Add-Content $outfile "$tab</$class>"
        Write-Output "[WARNING] There was an unexpected error while getting the $class class. Moving on to next section..."
        $warnings ++
        }
        #endregion GatherWMIVideoController

        #region GatherWMISoundDevice
        $class = "win32_sounddevice" # Multi item class
        try{
        Write-Output "[INFO] Getting details from class $class."
        $props = ('deviceid', 'manufacturer', 'name', 'pnpdeviceid', 'productname', 'statusinfo')
        $wmi = Get-WmiObject -Namespace root/cimv2 -Class $class -ComputerName $endpoint -ErrorAction Stop | Select-Object $props | Sort-Object name
        Add-Content $outfile "$tab<$class>"
            foreach($wm in $wmi){
            Add-Content $outfile "$tab$tab<$class`_multi>"
                Add-Content $outfile "$tab$tab$tab<deviceid>$($wm.deviceid)</deviceid>"
                Add-Content $outfile "$tab$tab$tab<manufacturer>$($wm.manufacturer)</manufacturer>"
                Add-Content $outfile "$tab$tab$tab<name>$($wm.name)</name>"
                Add-Content $outfile "$tab$tab$tab<pnpdeviceid>$($wm.pnpdeviceid)</pnpdeviceid>"
                Add-Content $outfile "$tab$tab$tab<productname>$($wm.productname)</productname>"
                Add-Content $outfile "$tab$tab$tab<statusinfo>$($wm.statusinfo)</statusinfo>"
            Add-Content $outfile "$tab$tab</$class`_multi>"
            }
        Add-Content $outfile "$tab</$class>"
        }
        catch{
        Add-Content $outfile "$tab<$class>"
        Add-Content $outfile "$tab$tab<errorcode>1</errorcode>"
        Add-Content $outfile "$tab$tab<errortext>$_.Exception.Message</errortext>"
        Add-Content $outfile "$tab</$class>"
        Write-Output "[WARNING] There was an unexpected error while getting the $class class. Moving on to next section..."
        $warnings ++
        }
        #endregion GatherWMISoundDevice

        #region GatherWMIPrinter
        $class = "win32_printer" # Multi item class
        try{
        Write-Output "[INFO] Getting details from class $class."
        $props = ('capabilitydescriptions', 'default', 'deviceid', 'drivername', 'hidden', 'local', 'name', 'network', 'pnpdeviceid', 'portname', 'printprocessor', 'servername', 'sharename', 'shared')
        $wmi = Get-WmiObject -Namespace root/cimv2 -Class $class -ComputerName $endpoint -ErrorAction Stop | Select-Object $props | Sort-Object name
        Add-Content $outfile "$tab<$class>"
            foreach($wm in $wmi){
            Add-Content $outfile "$tab$tab<$class`_multi>"
                Add-Content $outfile "$tab$tab$tab<capabilitydescriptions>$($wm.capabilitydescriptions)</capabilitydescriptions>"
                Add-Content $outfile "$tab$tab$tab<default>$($wm.default)</default>"
                Add-Content $outfile "$tab$tab$tab<deviceid>$($wm.deviceid)</deviceid>"
                Add-Content $outfile "$tab$tab$tab<drivername>$($wm.drivername)</drivername>"
                Add-Content $outfile "$tab$tab$tab<hidden>$($wm.hidden)</hidden>"
                Add-Content $outfile "$tab$tab$tab<local>$($wm.local)</local>"
                Add-Content $outfile "$tab$tab$tab<name>$($wm.name)</name>"
                Add-Content $outfile "$tab$tab$tab<network>$($wm.network)</network>"
                Add-Content $outfile "$tab$tab$tab<pnpdeviceid>$($wm.pnpdeviceid)</pnpdeviceid>"
                Add-Content $outfile "$tab$tab$tab<portname>$($wm.portname)</portname>"
                Add-Content $outfile "$tab$tab$tab<printprocessor>$($wm.printprocessor)</printprocessor>"
                Add-Content $outfile "$tab$tab$tab<servername>$($wm.servername)</servername>"
                Add-Content $outfile "$tab$tab$tab<sharename>$($wm.sharename)</sharename>"
                Add-Content $outfile "$tab$tab$tab<shared>$($wm.shared)</shared>"
            Add-Content $outfile "$tab$tab</$class`_multi>"
            }
        Add-Content $outfile "$tab</$class>"
        }
        catch{
        Add-Content $outfile "$tab<$class>"
        Add-Content $outfile "$tab$tab<errorcode>1</errorcode>"
        Add-Content $outfile "$tab$tab<errortext>$_.Exception.Message</errortext>"
        Add-Content $outfile "$tab</$class>"
        Write-Output "[WARNING] There was an unexpected error while getting the $class class. Moving on to next section..."
        $warnings ++
        }
        #endregion GatherWMIPrinter

        #region GatherWMIDiskDrive
        $class = "win32_diskdrive" # Multi item class
        try{
        Write-Output "[INFO] Getting details from class $class."
        $props = ('caption', 'firmwarerevision', 'interfacetype', 'manufacturer', 'model', 'name', 'pnpdeviceid', 'partitions', 'serialnumber', 'size')
        $wmi = Get-WmiObject -Namespace root/cimv2 -Class $class -ComputerName $endpoint -ErrorAction Stop | Select-Object $props | Sort-Object name
        Add-Content $outfile "$tab<$class>"
            foreach($wm in $wmi){
            Add-Content $outfile "$tab$tab<$class`_multi>"
                Add-Content $outfile "$tab$tab$tab<caption>$($wm.caption)</caption>"
                Add-Content $outfile "$tab$tab$tab<firmwarerevision>$($wm.firmwarerevision)</firmwarerevision>"
                Add-Content $outfile "$tab$tab$tab<interfacetype>$($wm.interfacetype)</interfacetype>"
                Add-Content $outfile "$tab$tab$tab<manufacturer>$($wm.manufacturer)</manufacturer>"
                Add-Content $outfile "$tab$tab$tab<model>$($wm.model)</model>"
                Add-Content $outfile "$tab$tab$tab<name>$($wm.name)</name>"
                Add-Content $outfile "$tab$tab$tab<pnpdeviceid>$($wm.pnpdeviceid)</pnpdeviceid>"
                Add-Content $outfile "$tab$tab$tab<partitions>$($wm.partitions)</partitions>"
                Add-Content $outfile "$tab$tab$tab<serialnumber>$($wm.serialnumber)</serialnumber>"
                Add-Content $outfile "$tab$tab$tab<size>$($wm.size)</size>"
            Add-Content $outfile "$tab$tab</$class`_multi>"
            }
        Add-Content $outfile "$tab</$class>"
        }
        catch{
        Add-Content $outfile "$tab<$class>"
        Add-Content $outfile "$tab$tab<errorcode>1</errorcode>"
        Add-Content $outfile "$tab$tab<errortext>$_.Exception.Message</errortext>"
        Add-Content $outfile "$tab</$class>"
        Write-Output "[WARNING] There was an unexpected error while getting the $class class. Moving on to next section..."
        $warnings ++
        }
        #endregion GatherWMIDiskDrive

        #region GatherWMIOptionalFeature
        $class = "win32_optionalfeature" # Multi item class
        try{
        Write-Output "[INFO] Getting details from class $class."
        $props = ('caption', 'name')
        $wmi = Get-WmiObject -Namespace root/cimv2 -Class $class -ComputerName $endpoint -ErrorAction Stop | Where-Object {$_.InstallState -eq 1} | Select-Object $props | Sort-Object name
        Add-Content $outfile "$tab<$class>"
            foreach($wm in $wmi){
            Add-Content $outfile "$tab$tab<$class`_multi>"
                Add-Content $outfile "$tab$tab$tab<caption>$($wm.caption)</caption>"
                Add-Content $outfile "$tab$tab$tab<name>$($wm.name)</name>"
            Add-Content $outfile "$tab$tab</$class`_multi>"
            }
        Add-Content $outfile "$tab</$class>"
        }
        catch{
        Add-Content $outfile "$tab<$class>"
        Add-Content $outfile "$tab$tab<errorcode>1</errorcode>"
        Add-Content $outfile "$tab$tab<errortext>$_.Exception.Message</errortext>"
        Add-Content $outfile "$tab</$class>"
        Write-Output "[WARNING] There was an unexpected error while getting the $class class. Moving on to next section..."
        $warnings ++
        }
        #endregion GatherWMIOptionalFeature

        #region GatherWMIEncryptableVolume
        $class = "win32_encryptablevolume" # Multi item class
        try{
        Write-Output "[INFO] Getting details from class $class."
        $props = ('conversionstatus', 'driveletter', 'encryptionmethod', 'isvolumeinitializedforprotection', 'protectionstatus', 'volumetype')
        $wmi = Get-WmiObject -Namespace root\CIMv2\Security\MicrosoftVolumeEncryption -Class $class -ComputerName $endpoint -ErrorAction Stop | Select-Object $props | Sort-Object driveletter
        Add-Content $outfile "$tab<$class>"
            foreach($wm in $wmi){
            Add-Content $outfile "$tab$tab<$class`_multi>"
                Add-Content $outfile "$tab$tab$tab<conversionstatus>$($wm.conversionstatus)</conversionstatus>"
                Add-Content $outfile "$tab$tab$tab<driveletter>$($wm.driveletter)</driveletter>"
                Add-Content $outfile "$tab$tab$tab<encryptionmethod>$($wm.encryptionmethod)</encryptionmethod>"
                Add-Content $outfile "$tab$tab$tab<isvolumeinitializedforprotection>$($wm.isvolumeinitializedforprotection)</isvolumeinitializedforprotection>"
                Add-Content $outfile "$tab$tab$tab<protectionstatus>$($wm.protectionstatus)</protectionstatus>"
                Add-Content $outfile "$tab$tab$tab<volumetype>$($wm.volumetype)</volumetype>"
            Add-Content $outfile "$tab$tab</$class`_multi>"
            }
        Add-Content $outfile "$tab</$class>"
        }
        catch{
        Add-Content $outfile "$tab<$class>"
        Add-Content $outfile "$tab$tab<errorcode>1</errorcode>"
        Add-Content $outfile "$tab$tab<errortext>$_.Exception.Message</errortext>"
        Add-Content $outfile "$tab</$class>"
        Write-Output "[WARNING] There was an unexpected error while getting the $class class. Moving on to next section..."
        $warnings ++
        }
        #endregion GatherWMIEncryptableVolume

        #region GatherWMIFirewallProduct
        $class = "firewallproduct" # Multi item class
        try{
        Write-Output "[INFO] Getting details from class $class."
        $props = ('displayname', 'pathtosignedproductexe', 'pathtosignedreportingexe', 'productstate', 'timestamp')
        $wmi = Get-WmiObject -Namespace root\SecurityCenter2 -Class $class -ComputerName $endpoint -ErrorAction Stop | Select-Object $props | Sort-Object displayname
        Add-Content $outfile "$tab<$class>"
            foreach($wm in $wmi){
            Add-Content $outfile "$tab$tab<$class`_multi>"
                Add-Content $outfile "$tab$tab$tab<displayname>$($wm.displayname)</displayname>"
                Add-Content $outfile "$tab$tab$tab<pathtosignedproductexe>$($wm.pathtosignedproductexe)</pathtosignedproductexe>"
                Add-Content $outfile "$tab$tab$tab<pathtosignedreportingexe>$($wm.pathtosignedreportingexe)</pathtosignedreportingexe>"
                Add-Content $outfile "$tab$tab$tab<productstate>$($wm.productstate)</productstate>"
                Add-Content $outfile "$tab$tab$tab<timestamp>$($wm.timestamp)</timestamp>"
            Add-Content $outfile "$tab$tab</$class`_multi>"
            }
        Add-Content $outfile "$tab</$class>"
        }
        catch{
        Add-Content $outfile "$tab<$class>"
        Add-Content $outfile "$tab$tab<errorcode>1</errorcode>"
        Add-Content $outfile "$tab$tab<errortext>$_.Exception.Message</errortext>"
        Add-Content $outfile "$tab</$class>"
        Write-Output "[WARNING] There was an unexpected error while getting the $class class. Moving on to next section..."
        $warnings ++
        }
        #endregion GatherWMIFirewallProduct

        #region GatherWMIAntiVirusProduct
        $class = "antivirusproduct" # Multi item class
        try{
        Write-Output "[INFO] Getting details from class $class."
        $props = ('displayname', 'pathtosignedproductexe', 'pathtosignedreportingexe', 'productstate', 'timestamp')
        $wmi = Get-WmiObject -Namespace root\SecurityCenter2 -Class $class -ComputerName $endpoint -ErrorAction Stop | Select-Object $props | Sort-Object displayname
        Add-Content $outfile "$tab<$class>"
            foreach($wm in $wmi){
            Add-Content $outfile "$tab$tab<$class`_multi>"
                Add-Content $outfile "$tab$tab$tab<displayname>$($wm.displayname)</displayname>"
                Add-Content $outfile "$tab$tab$tab<pathtosignedproductexe>$($wm.pathtosignedproductexe)</pathtosignedproductexe>"
                Add-Content $outfile "$tab$tab$tab<pathtosignedreportingexe>$($wm.pathtosignedreportingexe)</pathtosignedreportingexe>"
                Add-Content $outfile "$tab$tab$tab<productstate>$($wm.productstate)</productstate>"
                Add-Content $outfile "$tab$tab$tab<timestamp>$($wm.timestamp)</timestamp>"
            Add-Content $outfile "$tab$tab</$class`_multi>"
            }
        Add-Content $outfile "$tab</$class>"
        }
        catch{
        Add-Content $outfile "$tab<$class>"
        Add-Content $outfile "$tab$tab<errorcode>1</errorcode>"
        Add-Content $outfile "$tab$tab<errortext>$_.Exception.Message</errortext>"
        Add-Content $outfile "$tab</$class>"
        Write-Output "[WARNING] There was an unexpected error while getting the $class class. Moving on to next section..."
        $warnings ++
        }
        #endregion GatherWMIAntiVirusProduct

        #region GatherWMIAntiSpywareProduct
        $class = "antispywareproduct" # Multi item class
        try{
        Write-Output "[INFO] Getting details from class $class."
        $props = ('displayname', 'pathtosignedproductexe', 'pathtosignedreportingexe', 'productstate', 'timestamp')
        $wmi = Get-WmiObject -Namespace root\SecurityCenter2 -Class $class -ComputerName $endpoint -ErrorAction Stop | Select-Object $props | Sort-Object displayname
        Add-Content $outfile "$tab<$class>"
            foreach($wm in $wmi){
            Add-Content $outfile "$tab$tab<$class`_multi>"
                Add-Content $outfile "$tab$tab$tab<displayname>$($wm.displayname)</displayname>"
                Add-Content $outfile "$tab$tab$tab<pathtosignedproductexe>$($wm.pathtosignedproductexe)</pathtosignedproductexe>"
                Add-Content $outfile "$tab$tab$tab<pathtosignedreportingexe>$($wm.pathtosignedreportingexe)</pathtosignedreportingexe>"
                Add-Content $outfile "$tab$tab$tab<productstate>$($wm.productstate)</productstate>"
                Add-Content $outfile "$tab$tab$tab<timestamp>$($wm.timestamp)</timestamp>"
            Add-Content $outfile "$tab$tab</$class`_multi>"
            }
        Add-Content $outfile "$tab</$class>"
        }
        catch{
        Add-Content $outfile "$tab<$class>"
        Add-Content $outfile "$tab$tab<errorcode>1</errorcode>"
        Add-Content $outfile "$tab$tab<errortext>$_.Exception.Message</errortext>"
        Add-Content $outfile "$tab</$class>"
        Write-Output "[WARNING] There was an unexpected error while getting the $class class. Moving on to next section..."
        $warnings ++
        }
        #endregion GatherWMIAntiVirusProduct

        #region GatherRegistryInstalledPrograms
        try{
        Write-Output "[INFO] Getting installed program details from the registry."
        $reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $endpoint)
        $key = "SYSTEM\\CurrentControlSet\\Control\\Session Manager\\Environment"
        $openSubKey = $reg.OpenSubKey($key)
        $arch = $openSubKey.getvalue("PROCESSOR_ARCHITECTURE")
        Add-Content $outfile "$tab<installed_programs>"

            # 64 bit architecture detected
            if($arch -eq 'AMD64'){

            Write-Output "[INFO] Getting installed programs from registry (x64)."
            $key = "SOFTWARE\\WOW6432Node\\Microsoft\\Windows\\CurrentVersion\\Uninstall"
            $openSubKey = $reg.OpenSubKey($key)
            $subKeys = $openSubKey.GetSubKeyNames()

               foreach ($subKey in $subKeys){

                # replace , with ; for any possible values that could have
                $dispName = ($reg.OpenSubKey($key+"\\"+$subKey).getValue('DisplayName')) -replace ",", ";"
                $sysComp = ($reg.OpenSubKey($key+"\\"+$subKey).getValue('SystemComponent'))
                $parKeyName = ($reg.OpenSubKey($key+"\\"+$subKey).getValue('ParentKeyName'))
                $dispVer = ($reg.OpenSubKey($key+"\\"+$subKey).getValue('DisplayVersion')) -replace ",", ";"
                $publisher = ($reg.OpenSubKey($key+"\\"+$subKey).getValue('Publisher')) -replace ",", ";"
                $comments = ($reg.OpenSubKey($key+"\\"+$subKey).getValue('Comments')) -replace ",",";"

                    # display name is blank and system components is not 1 and partent key name is blank
                    if(($dispName.length -ne 0) -and ($sysComp -ne 1) -and ($parKeyName.length -eq 0)){
                    Add-Content $outfile "$tab$tab<installed_programs_multi>"
                    Add-Content $outfile "$tab$tab$tab<source>System</source>"
                    Add-Content $outfile "$tab$tab$tab<name>$dispName</name>"
                    Add-Content $outfile "$tab$tab$tab<version>$dispVer</version>"
                    Add-Content $outfile "$tab$tab$tab<location>$($reg.OpenSubKey($key+"\\"+$subKey).getValue('InstallLocation'))</location>"
                    Add-Content $outfile "$tab$tab$tab<publisher>$publisher</publisher>"
                    Add-Content $outfile "$tab$tab$tab<install_date>$($reg.OpenSubKey($key+"\\"+$subKey).getValue('InstallDate'))</install_date>"
                    Add-Content $outfile "$tab$tab$tab<architecture>x64</architecture>"
                    Add-Content $outfile "$tab$tab$tab<comments>$comments</comments>"
                    Add-Content $outfile "$tab$tab</installed_programs_multi>"
                    }

                }

            }

        # x86 Programs. There will always be x86 programs to get even if the OS is x64
        Write-Output "[INFO] Getting installed programs from registry (x86)."
        $key = "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall"
        $openSubKey = $reg.OpenSubKey($key)
        $subKeys = $openSubKey.GetSubKeyNames()

            foreach ($subKey in $subKeys){

            # replace , with ; for any possible values that could have
            $dispName = ($reg.OpenSubKey($key+"\\"+$subKey).getValue('DisplayName')) -replace ",", ";"
            $sysComp = ($reg.OpenSubKey($key+"\\"+$subKey).getValue('SystemComponent'))
            $parKeyName = ($reg.OpenSubKey($key+"\\"+$subKey).getValue('ParentKeyName'))
            $dispVer = ($reg.OpenSubKey($key+"\\"+$subKey).getValue('DisplayVersion')) -replace ",", ";"
            $publisher = ($reg.OpenSubKey($key+"\\"+$subKey).getValue('Publisher')) -replace ",", ";"
            $comments = ($reg.OpenSubKey($key+"\\"+$subKey).getValue('Comments')) -replace ",",";"

                # display name is blank and system components is not 1 and partent key name is blank
                if(($dispName.length -ne 0) -and ($sysComp -ne 1) -and ($parKeyName.length -eq 0)){
                Add-Content $outfile "$tab$tab<installed_programs_multi>"
                Add-Content $outfile "$tab$tab$tab<source>System</source>"
                Add-Content $outfile "$tab$tab$tab<name>$dispName</name>"
                Add-Content $outfile "$tab$tab$tab<version>$dispVer</version>"
                Add-Content $outfile "$tab$tab$tab<location>$($reg.OpenSubKey($key+"\\"+$subKey).getValue('InstallLocation'))</location>"
                Add-Content $outfile "$tab$tab$tab<publisher>$publisher</publisher>"
                Add-Content $outfile "$tab$tab$tab<install_date>$($reg.OpenSubKey($key+"\\"+$subKey).getValue('InstallDate'))</install_date>"
                Add-Content $outfile "$tab$tab$tab<architecture>x86</architecture>"
                Add-Content $outfile "$tab$tab$tab<comments>$comments</comments>"
                Add-Content $outfile "$tab$tab</installed_programs_multi>"
                }

            }

            # Only run if running locally
            if($remote -eq 'No'){
            # Current user programs - not listed in locations above
            Write-Output "[INFO] Getting installed programs from the current user ($env:USERDOMAIN\$env:USERNAME)."
            $reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('CurrentUser', $endpoint)
            $key = "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall"
            $openSubKey = $reg.OpenSubKey($key)
            $subKeys = $openSubKey.GetSubKeyNames()

                foreach ($subKey in $subKeys){

                # replace , with ; for any possible values that could have
                $dispName = ($reg.OpenSubKey($key+"\\"+$subKey).getValue('DisplayName')) -replace ",", ";"
                $sysComp = ($reg.OpenSubKey($key+"\\"+$subKey).getValue('SystemComponent'))
                $parKeyName = ($reg.OpenSubKey($key+"\\"+$subKey).getValue('ParentKeyName'))
                $dispVer = ($reg.OpenSubKey($key+"\\"+$subKey).getValue('DisplayVersion')) -replace ",", ";"
                $publisher = ($reg.OpenSubKey($key+"\\"+$subKey).getValue('Publisher')) -replace ",", ";"
                $comments = ($reg.OpenSubKey($key+"\\"+$subKey).getValue('Comments')) -replace ",",";"

                    # display name is blank and system components is not 1 and partent key name is blank
                    if(($dispName.length -ne 0) -and ($sysComp -ne 1) -and ($parKeyName.length -eq 0)){
                    Add-Content $outfile "$tab$tab<installed_programs_multi>"
                    Add-Content $outfile "$tab$tab$tab<source>$env:userdomain\$env:username</source>"
                    Add-Content $outfile "$tab$tab$tab<name>$dispName</name>"
                    Add-Content $outfile "$tab$tab$tab<version>$dispVer</version>"
                    Add-Content $outfile "$tab$tab$tab<location>$($reg.OpenSubKey($key+"\\"+$subKey).getValue('InstallLocation'))</location>"
                    Add-Content $outfile "$tab$tab$tab<publisher>$publisher</publisher>"
                    Add-Content $outfile "$tab$tab$tab<install_date>$($reg.OpenSubKey($key+"\\"+$subKey).getValue('InstallDate'))</install_date>"
                    Add-Content $outfile "$tab$tab$tab<architecture>x86</architecture>"
                    Add-Content $outfile "$tab$tab$tab<comments>$comments</comments>"
                    Add-Content $outfile "$tab$tab</installed_programs_multi>"
                    }

                }

            }

        Add-Content $outfile "$tab</installed_programs>"

        }
        catch{
        Add-Content $outfile "$tab<installed_programs>"
        Add-Content $outfile "$tab$tab<errorcode>1</errorcode>"
        Add-Content $outfile "$tab$tab<errortext>$_.Exception.Message</errortext>"
        Add-Content $outfile "$tab</installed_programs>"
        Write-Output "[WARNING] There was an unexpected error while getting the installed program details. Moving on to next section..."
        $warnings ++
        }
        #endregion GatherRegistryInstalledPrograms

        #region GatherComputerCertificates
        try{
        Write-Output "[INFO] Getting computer certificates from certificates store."
        # Get all stores in use from the registry
        $reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $endpoint)
        $key = "SOFTWARE\\Microsoft\\SystemCertificates"
        $openSubKey = $reg.OpenSubKey($key)
        $subKeys = $openSubKey.GetSubKeyNames()

        $cryptOpenFlags = [System.Security.Cryptography.X509Certificates.OpenFlags]"ReadOnly"
        $cryptStoreLocation=[System.Security.Cryptography.X509Certificates.StoreLocation]"LocalMachine"

        Add-Content $outfile "$tab<computer_certificates>"

            foreach ($sub in $subKeys){

            # Open each store as identified in the registry
            $store = New-Object System.Security.Cryptography.X509Certificates.X509Store("\\$endpoint\$sub",$cryptStoreLocation)
            $store.Open($cryptOpenFlags)
            $storeCerts = $store.Certificates

                foreach($storeCert in $storeCerts){

                # Get content up until 1st ',' then replace CN, OU etc...
                $storeSubjectCleaned = (((($storeCert.Subject).split(',')[0] -replace "CN=", "") -replace "OU=", "") -replace "E=", "") -replace "`"", ""

                    # Get content up until 1st ',' then replace CN, OU etc...
                    if (($storeCert.Issuer).Length -ne 0){
                    $storeIssuerCleaned = (((($storeCert.Issuer).split(',')[0] -replace "CN=", "") -replace "OU=", "") -replace "E=", "") -replace "`"", ""
                    }
                    else{
                    $storeIssuerCleaned = ""
                    }

                Add-Content $outfile "$tab$tab<computer_certificates_multi>"
                Add-Content $outfile "$tab$tab$tab<store>$sub</store>"
                Add-Content $outfile "$tab$tab$tab<subject>$storeSubjectCleaned</subject>"
                Add-Content $outfile "$tab$tab$tab<issuer>$storeIssuerCleaned</issuer>"
                Add-Content $outfile "$tab$tab$tab<validfrom>$($storeCert.GetEffectiveDateString())</validfrom>"
                Add-Content $outfile "$tab$tab$tab<expiration>$($storeCert.GetExpirationDateString())</expiration>"
                Add-Content $outfile "$tab$tab$tab<thumbprint>$($storeCert.Thumbprint)</thumbprint>"
                Add-Content $outfile "$tab$tab$tab<serialnumber>$($storeCert.SerialNumber)</serialnumber>"
                Add-Content $outfile "$tab$tab$tab<format>$($storeCert.GetFormat())</format>"
                Add-Content $outfile "$tab$tab$tab<version>$($storeCert.Version)</version>"
                Add-Content $outfile "$tab$tab$tab<signaturealgorithmfriendlyname>$($storeCert.SignatureAlgorithm.FriendlyName)</signaturealgorithmfriendlyname>"
                Add-Content $outfile "$tab$tab$tab<signaturealgorithmvalue>$($storeCert.SignatureAlgorithm.Value)</signaturealgorithmvalue>"
                Add-Content $outfile "$tab$tab$tab<enhancedkeyusagelistfriendlyname>$($storeCert.EnhancedKeyUsageList.FriendlyName)</enhancedkeyusagelistfriendlyname>"
                Add-Content $outfile "$tab$tab$tab<archived>$($storeCert.Archived)</archived>"
                Add-Content $outfile "$tab$tab$tab<hasprivatekey>$($storeCert.HasPrivateKey)</hasprivatekey>"
                Add-Content $outfile "$tab$tab$tab<friendlyname>$($storeCert.FriendlyName)</friendlyname>"
                Add-Content $outfile "$tab$tab</computer_certificates_multi>"

                }

            }

        Add-Content $outfile "$tab</computer_certificates>"

        }
        catch{
        Add-Content $outfile "$tab<computer_certificates>"
        Add-Content $outfile "$tab$tab<errorcode>1</errorcode>"
        Add-Content $outfile "$tab$tab<errortext>$_.Exception.Message</errortext>"
        Add-Content $outfile "$tab</computer_certificates>"
        Write-Output "[WARNING] There was an unexpected error while getting the computer certificate details. Moving on to next section..."
        $warnings ++
        }
        #endregion GatherComputerCertificates

       #region GatherWMIPowerPlan
        $class = "win32_powerplan" # Multi item class
        try{
        Write-Output "[INFO] Getting details from class $class."
        $props = ('elementname', 'description', 'instanceid', 'isactive')
        $wmi = Get-WmiObject -Namespace root/cimv2/power -Class $class -ComputerName $endpoint -ErrorAction Stop | Select-Object $props | Sort-Object elementname
        Add-Content $outfile "$tab<$class>"
            foreach($wm in $wmi){
            Add-Content $outfile "$tab$tab<$class`_multi>"
                Add-Content $outfile "$tab$tab$tab<elementname>$($wm.elementname)</elementname>"
                Add-Content $outfile "$tab$tab$tab<description>$($wm.description)</description>"
                Add-Content $outfile "$tab$tab$tab<instanceid>$($wm.instanceid)</instanceid>"
                Add-Content $outfile "$tab$tab$tab<isactive>$($wm.isactive)</isactive>"
            Add-Content $outfile "$tab$tab</$class`_multi>"
            }
        Add-Content $outfile "$tab</$class>"
        }
        catch{
        Add-Content $outfile "$tab<$class>"
        Add-Content $outfile "$tab$tab<errorcode>1</errorcode>"
        Add-Content $outfile "$tab$tab<errortext>$_.Exception.Message</errortext>"
        Add-Content $outfile "$tab</$class>"
        Write-Output "[WARNING] There was an unexpected error while getting the $class class. Moving on to next section..."
        $warnings ++
        }
        #endregion GatherWMIPowerPlan

    #region GatherEnd
    # Close end tag
    Add-Content $outfile "</info>"

        ### CONVERT TO TRUE XML ###
        try{
        Write-Output "[INFO] Converting $outFile to true xml."
        (Get-Content $outFile -ErrorAction SilentlyContinue).replace('&', '&amp;') | Set-Content $outFile -ErrorAction SilentlyContinue
        }
        catch{
        Write-Output "[ERROR] There was an unexpected error while converting $outFile to true xml. Script terminated."
        Write-Output "[ERROR] $($_.Exception.Message)."
        break
        }

    } # end if -ne to reportmode
    #endregion GatherEnd
#endregion Gather

#region Report
    #region ReportInitialize
    if($mode -eq "ReportOnly"){

        # Verify xml report file exists
        if(!(Test-Path $reportFile -Include *.xml)){
        Write-Output "[ERROR] The report file $reportFile was not found. Script terminated."
        break
        }
        else{
        Write-Output "[INFO] Report file $reportFile selected."
        }
    }

    if($mode -ne "GatherOnly"){

    Write-Output "[INFO] Generating $reportType report."

    ### GET FILE ###
        try{
        Write-Output "[INFO] Reading contents of $reportFile."
        [xml]$xml = Get-Content -Path $reportFile -ErrorAction Stop
        }
        catch{
        Write-Output "[ERROR] Unable to open $reportFile. Script terminated!"
        Write-Output "[ERROR] $($_.exception.message)"
        break
        }

    ### FORMAT INVENT DATE ###
    $inventDate = ($xml.info.general.ge_inventdate).Substring(6,2) + "/" + ($xml.info.general.ge_inventdate).Substring(4,2) + "/" + `
    ($xml.info.general.ge_inventdate).Substring(0,4) + " " + ($xml.info.general.ge_inventdate).Substring(8,2) + ":" + ($xml.info.general.ge_inventdate).Substring(10,2)
    #endregion ReportInitialize

    #region ReportCreateWordDoc
        try{
        Write-Output "[INFO] Constructing Word document."
        $word = New-Object -ComObject Word.Application -ErrorAction SilentlyContinue
        }
        catch{
        Write-Output "[ERROR] Unable to construct the Word document. Check that Word is installed. Script terminated!"
        Write-Output "[ERROR] $($_.exception.message)"

            if($mode -eq "GatherAndReport"){
            Write-Output "[ERROR] Note: $reportFile has been generated. Rerun script in ReportOnly mode on a endpoint with Word installed."
            }

        break
        }

    $word.Visible = $False
    $document = $word.Documents.Add()
    $section = $document.Sections.Item(1);
    $header = $section.Headers.Item(1);
    $header.Range.Text = "https://github.com/gordonrankine/get-win10info";
    [void]$document.Sections.Item(1).Footers.Item(3).Pagenumbers.Add()
    $selection = $word.Selection

    ### TITLE ###
    $selection.Style = 'Title'
    $selection.TypeText("Build Report for $($xml.info.general.ge_server)")
    $selection.TypeParagraph()
    $selection.TypeParagraph()
    $selection.TypeParagraph()

    $selection.Style = 'SubTitle'
    $selection.Font.Bold = 1
    $selection.TypeText("Report type: ")
    $selection.Font.Bold = 0
    $selection.TypeText("$reportType")
    $selection.TypeParagraph()

    $selection.Style = 'SubTitle'
    $selection.Font.Bold = 1
    $selection.TypeText("Inventory date: ")
    $selection.Font.Bold = 0
    $selection.TypeText("$inventDate")
    $selection.TypeParagraph()

    $selection.Style = 'SubTitle'
    $selection.Font.Bold = 1
    $selection.TypeText("Inventory script: ")
    $selection.Font.Bold = 0
    $selection.TypeText("$($xml.info.general.ge_scriptname)")
    $selection.TypeParagraph()

    $selection.Style = 'SubTitle'
    $selection.Font.Bold = 1
    $selection.TypeText("Run by: ")
    $selection.Font.Bold = 0
    $selection.TypeText("$($xml.info.general.ge_runby)")
    $selection.InsertNewPage()

    ### TABLE OF CONTENTS ###
    $range = $selection.Range
    $toc = $document.TablesOfContents.Add($range)
    $selection.InsertNewPage()
    #endregion ReportCreateWordDoc

    #region ReportOperatingSystem
    # BASIC
    $selection.style = 'Heading 1'
    $selection.TypeText("[Registry] Operating System")
    $selection.TypeParagraph()
    $selection.TypeParagraph()

    $selection.Style = 'Normal'
    $selection.TypeText("This data is collected from the registry key (HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion)")
    $selection.TypeParagraph()
    $selection.TypeParagraph()

        if(($xml.info.winver).count -eq 0){
        Write-Output "[WARNING] Registry Operating System details not found. Moving on to next section..."
        $warnings ++
        $selection.Style = 'Normal'
        $selection.Font.Color="255"
        $selection.TypeText("[WARNING] Registry Operating System data not found!")
        $selection.TypeParagraph()
        }
        elseif(($xml.info.winver.errorcode) -eq 1){
        Write-Output "[WARNING] Registry Operating System details not collected. Moving on to next section..."
        $warnings ++
        $selection.Style = 'Normal'
        $selection.Font.Color="255"
        $selection.TypeText("[WARNING] Registry Operating System data not collected!")
        $selection.TypeParagraph()
        $selection.TypeText("Reason for error: $($xml.info.winver.errortext)")
        $selection.TypeParagraph()
        }
        else{
        Write-Output "[INFO] Populating Registry Operating System table."

        $table = $selection.Tables.add(
        $selection.Range,
        8,
        2,
        [Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior,
        [Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitWindow
        )

        $table.style = "Grid Table 4 - Accent 1"
        $table.cell(1,1).range.text = "Item"
        $table.cell(1,2).range.text = "Value"

        $table.cell(2,1).range.text = "Product Name"
        $table.cell(2,2).range.text = $xml.info.winver.productname
        $table.cell(3,1).range.text = "Edition ID"
        $table.cell(3,2).range.text = $xml.info.winver.editionid
        $table.cell(4,1).range.text = "Build"
        $table.cell(4,2).range.text = $xml.info.winver.build
        $table.cell(5,1).range.text = "Build Branch"
        $table.cell(5,2).range.text = $xml.info.winver.buildbranch
        $table.cell(6,1).range.text = "Version"
        $table.cell(6,2).range.text = $xml.info.winver.version
        $table.cell(7,1).range.text = "Registered Organisation"
        $table.cell(7,2).range.text = $xml.info.winver.registeredorganization
        $table.cell(8,1).range.text = "Registered Owner"
        $table.cell(8,2).range.text = $xml.info.winver.registeredowner

        $table.Rows.item(1).Headingformat=-1
        $table.ApplyStyleFirstColumn = $false
        $selection.InsertCaption(-2, ": [Registry] Operating System", $null, 1, $false)

        }

    $selection.EndOf(15) | Out-Null
    $selection.MoveDown() | Out-Null
    $selection.InsertNewPage()
    #endregion ReportOperatingSystem

    #region ReportInstalledPrograms
    #BASIC & DETAILED
    $selection.style = 'Heading 1'
    $selection.TypeText("[Registry] Installed Programs (System)")
    $selection.TypeParagraph()
    $selection.TypeParagraph()

    $selection.Style = 'Normal'
    $selection.TypeText("This data is collected from the Windows Registry subkeys HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall & HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall")
    $selection.TypeParagraph()
    $selection.TypeText("Note: Programs installed for all users are shown here.")
    $selection.TypeParagraph()
    $selection.TypeParagraph()

        if($reportType -ne 'Basic'){
        $selection.style = 'Heading 2'
        $selection.TypeText("Basic")
        $selection.TypeParagraph()
        $selection.TypeParagraph()
        }

        if(($xml.info.installed_programs).count -eq 0){
        Write-Output "[WARNING] Registry Installed Programs - System (Basic) details not found. Moving on to next section..."
        $warnings ++
        $selection.Style = 'Normal'
        $selection.Font.Color="255"
        $selection.TypeText("[WARNING] Registry Installed Programs - System (Basic) data not found!")
        $selection.TypeParagraph()
        }
        elseif(($xml.info.installed_programs.errorcode) -eq 1){
        Write-Output "[WARNING] Registry Installed Programs - System (Basic) details not collected. Moving on to next section..."
        $warnings ++
        $selection.Style = 'Normal'
        $selection.Font.Color="255"
        $selection.TypeText("[WARNING] Registry Installed Programs - System (Basic) data not collected!")
        $selection.TypeParagraph()
        $selection.TypeText("Reason for error: $($xml.info.installed_programs.errortext)")
        $selection.TypeParagraph()
        }
        else{
        Write-Output "[INFO] Populating Registry Installed Programs - System (Basic) table."

        # Count rows in multi
        $multis = $xml.selectnodes("//info/installed_programs/installed_programs_multi") | Where-Object {$_.source -eq 'System'} | Sort-Object {$_.name}
        $count = ($multis | Measure-Object).Count

            # Calculate rows (from query above) and add 1 for the header or 2 if there is no data
            if($count -eq 0){
            $rows = $count + 2
            }
            else{
            $rows = $count + 1
            }

        $table = $selection.Tables.add(
        $selection.Range,
        $rows,
        4,
        [Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior,
        [Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitWindow
        )

        $table.style = "Grid Table 4 - Accent 1"
        $table.cell(1,1).range.text = "Name"
        $table.cell(1,2).range.text = "Publisher"
        $table.cell(1,3).range.text = "Installed On"
        $table.cell(1,4).range.text = "Version"

        $i = 2 # 2 as header is row 1

            if($count -eq 0){
            $table.cell($i,1).range.text = "No programs installed"
            $table.Cell($i,2).Merge($table.Cell($i,1))
            $table.Cell($i,2).Merge($table.Cell($i,1))
            $table.Cell($i,2).Merge($table.Cell($i,1))
            }

            foreach($multi in $multis){

            $table.cell($i,1).range.text = $multi.name
            $table.cell($i,2).range.text = $multi.publisher

                if($multi.install_date -ne ''){
                $installDate =
                $multi.install_date.Substring(6,2) + "/" +
                $multi.install_date.Substring(4,2) + "/" +
                $multi.install_date.Substring(0,4)
                }
                else{
                $installDate = $multi.install_date
                }

            $table.cell($i,3).range.text = $installDate
            $table.cell($i,4).range.text = $multi.version
            $i++

            }

        $table.Rows.item(1).Headingformat=-1
        $table.ApplyStyleFirstColumn = $false
        $selection.InsertCaption(-2, ": [Registry] Installed Programs - System (Basic)", $null, 1, $false)

          if($reportType -ne "Basic"){
          $selection.TypeParagraph()
          }

        }

    $selection.EndOf(15) | Out-Null
    $selection.MoveDown() | Out-Null

        if($reportType -eq "Detailed"){
        $selection.TypeParagraph()
        $selection.style = 'Heading 2'
        $selection.TypeText("Detailed")
        $selection.TypeParagraph()
        $selection.TypeParagraph()

            if(($xml.info.installed_programs).count -eq 0){
            Write-Output "[WARNING] Registry Installed Programs - System (Detailed) details not found. Moving on to next section..."
            $warnings ++
            $selection.Style = 'Normal'
            $selection.Font.Color="255"
            $selection.TypeText("[WARNING] Registry Installed Programs - System (Detailed) data not found!")
            $selection.TypeParagraph()
            }
            elseif(($xml.info.installed_programs.errorcode) -eq 1){
            Write-Output "[WARNING] Registry Installed Programs - System (Detailed) details not collected. Moving on to next section..."
            $warnings ++
            $selection.Style = 'Normal'
            $selection.Font.Color="255"
            $selection.TypeText("[WARNING] Registry Installed Programs - System (Detailed) data not collected!")
            $selection.TypeParagraph()
            $selection.TypeText("Reason for error: $($xml.info.installed_programs.errortext)")
            $selection.TypeParagraph()
            }
            else{
            Write-Output "[INFO] Populating Registry Installed Programs - System (Detailed) table."

            # Count rows in multi
            $multis = $xml.selectnodes("//info/installed_programs/installed_programs_multi") | Where-Object {$_.source -eq 'System'} | Sort-Object {$_.name}
            $count = ($multis | Measure-Object).Count

                # Calculate rows. (No of items x no of multis) + (no of multis + 1 (for header)) | or less 1 for header if single row
                if($count -eq 0){
                $rows = $count + 2
                }
                elseif($count -eq 1){
                $rows = (8 + $count)
                }
                else{
                $rows = (8 * $count) + ($count + 1)
                }

            $table = $selection.Tables.add(
            $selection.Range,
            $rows,
            2,
            [Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior,
            [Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitWindow
            )

            $table.style = "Grid Table 4 - Accent 1"
            $table.cell(1,1).range.text = "Item"
            $table.cell(1,2).range.text = "Value"

            $i = 2 # 2 as header is row 1
            $y = 1

                if($count -eq 0){
                $table.cell($i,1).range.text = "No programs installed"
                $table.Cell($i,2).Merge($table.Cell($i,1))
                }

                foreach($multi in $multis){

                    if($multi.source -eq 'System'){

                        if($count -gt 1){
                        $table.cell($i,1).Merge($table.cell($i, 2))
                        $table.cell($i,1).range.text = "------ BLOCK $y ------"
                        $i++
                        }

                    $table.cell($i,1).range.text = "Source"
                    $table.cell($i,2).range.text = $multi.source
                    $i++

                    $table.cell($i,1).range.text = "Name"
                    $table.cell($i,2).range.text = $multi.name
                    $i++

                    $table.cell($i,1).range.text = "Version"
                    $table.cell($i,2).range.text = $multi.version
                    $i++

                    $table.cell($i,1).range.text = "Location"
                    $table.cell($i,2).range.text = $multi.location
                    $i++

                    $table.cell($i,1).range.text = "Publisher"
                    $table.cell($i,2).range.text = $multi.publisher
                    $i++

                    $table.cell($i,1).range.text = "Install date"

                        if($multi.install_date -ne ''){
                        $installDate =
                        $multi.install_date.Substring(6,2) + "/" +
                        $multi.install_date.Substring(4,2) + "/" +
                        $multi.install_date.Substring(0,4)
                        }
                        else{
                        $installDate = $multi.install_date
                        }

                    $table.cell($i,2).range.text = $multi.installdate
                    $i++

                    $table.cell($i,1).range.text = "Architecture"
                    $table.cell($i,2).range.text = $multi.architecture
                    $i++

                    $table.cell($i,1).range.text = "Comments"
                    $table.cell($i,2).range.text = $multi.comments
                    $i++

                    $y++

                    }

                }

            $table.Rows.item(1).Headingformat=-1
            $table.ApplyStyleFirstColumn = $false
            $selection.InsertCaption(-2, ": [Registry] Installed Program - System (Detailed)", $null, 1, $false)

            }

        }

    $selection.EndOf(15) | Out-Null
    $selection.MoveDown() | Out-Null
    $selection.InsertNewPage()

        if($xml.info.general.ge_remote -eq 'No'){
        ## REGISTRY - INSTALLED PROGRAMS - USER - BASIC & DETAILED ###
        $selection.style = 'Heading 1'
        $selection.TypeText("[Registry] Installed Programs (User)")
        $selection.TypeParagraph()
        $selection.TypeParagraph()

        $selection.Style = 'Normal'
        $selection.TypeText("This data is collected from the Windows Registry subkeys HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall only from user `"$($multi.source)`".")
        $selection.TypeParagraph()
        $selection.TypeText("Note: This is data is included so that it can be used to cross reference what is seen in the `"Programs and Features`" applet. The installed programs from the user account running the script will be shown.")
        $selection.TypeParagraph()
        $selection.TypeParagraph()

            if($reportType -ne 'Basic'){
            $selection.style = 'Heading 2'
            $selection.TypeText("Basic")
            $selection.TypeParagraph()
            $selection.TypeParagraph()
            }

            if(($xml.info.installed_programs).count -eq 0){
            Write-Output "[WARNING] Registry Installed Programs - User (Basic) details not found. Moving on to next section..."
            $warnings ++
            $selection.Style = 'Normal'
            $selection.Font.Color="255"
            $selection.TypeText("[WARNING] Registry Installed Programs - User (Basic) data not found!")
            $selection.TypeParagraph()
            }
            elseif(($xml.info.installed_programs.errorcode) -eq 1){
            Write-Output "[WARNING] Registry Installed Programs - User (Basic) details not collected. Moving on to next section..."
            $warnings ++
            $selection.Style = 'Normal'
            $selection.Font.Color="255"
            $selection.TypeText("[WARNING] Registry Installed Programs - User (Basic) data not collected!")
            $selection.TypeParagraph()
            $selection.TypeText("Reason for error: $($xml.info.installed_programs.errortext)")
            $selection.TypeParagraph()
            }
            else{
            Write-Output "[INFO] Populating Registry Installed Programs - User (Basic) table."

            # Count rows in multi
            $multis = $xml.selectnodes("//info/installed_programs/installed_programs_multi") | Where-Object {$_.source -ne 'System'} | Sort-Object {$_.name}
            $count = ($multis | Measure-Object).Count

                # Calculate rows (from query above) and add 1 for the header
                if($count -eq 0){
                $rows = $count + 2
                }
                else{
                $rows = $count + 1
                }

            $table = $selection.Tables.add(
            $selection.Range,
            $rows,
            4,
            [Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior,
            [Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitWindow
            )

            $table.style = "Grid Table 4 - Accent 1"
            $table.cell(1,1).range.text = "Name"
            $table.cell(1,2).range.text = "Publisher"
            $table.cell(1,3).range.text = "Installed On"
            $table.cell(1,4).range.text = "Version"

            $i = 2 # 2 as header is row 1

                if($count -eq 0){
                $table.cell($i,1).range.text = "No programs installed"
                $table.Cell($i,2).Merge($table.Cell($i,1))
                $table.Cell($i,2).Merge($table.Cell($i,1))
                $table.Cell($i,2).Merge($table.Cell($i,1))
                }

                foreach($multi in $multis){

                $table.cell($i,1).range.text = $multi.name
                $table.cell($i,2).range.text = $multi.publisher

                    if($multi.install_date -ne ''){
                    $installDate =
                    $multi.install_date.Substring(6,2) + "/" +
                    $multi.install_date.Substring(4,2) + "/" +
                    $multi.install_date.Substring(0,4)
                    }
                    else{
                    $installDate = $multi.install_date
                    }

                $table.cell($i,3).range.text = $installDate
                $table.cell($i,4).range.text = $multi.version
                $i++

                }

            $table.Rows.item(1).Headingformat=-1
            $table.ApplyStyleFirstColumn = $false
            $selection.InsertCaption(-2, ": [Registry] Installed Programs - User (Basic)", $null, 1, $false)

              if($reportType -ne "Basic"){
              $selection.TypeParagraph()
              }

            }

        $selection.EndOf(15) | Out-Null
        $selection.MoveDown() | Out-Null

            if($reportType -eq "Detailed"){
            $selection.TypeParagraph()
            $selection.style = 'Heading 2'
            $selection.TypeText("Detailed")
            $selection.TypeParagraph()
            $selection.TypeParagraph()

                if(($xml.info.installed_programs).count -eq 0){
                Write-Output "[WARNING] Registry Installed Programs - User (Detailed) details not found. Moving on to next section..."
                $warnings ++
                $selection.Style = 'Normal'
                $selection.Font.Color="255"
                $selection.TypeText("[WARNING] Registry Installed Programs - User (Detailed) data not found!")
                $selection.TypeParagraph()
                }
                elseif(($xml.info.installed_programs.errorcode) -eq 1){
                Write-Output "[WARNING] Registry Installed Programs - User (Detailed) details not collected. Moving on to next section..."
                $warnings ++
                $selection.Style = 'Normal'
                $selection.Font.Color="255"
                $selection.TypeText("[WARNING] Registry Installed Programs - User (Detailed) data not collected!")
                $selection.TypeParagraph()
                $selection.TypeText("Reason for error: $($xml.info.installed_programs.errortext)")
                $selection.TypeParagraph()
                }
                else{
                Write-Output "[INFO] Populating Registry Installed Programs - User (Detailed) table."

                # Count rows in multi
                $multis = $xml.selectnodes("//info/installed_programs/installed_programs_multi") | Where-Object {$_.source -ne 'System'} | Sort-Object {$_.name}
                $count = ($multis | Measure-Object).Count

                    # Calculate rows. (No of items x no of multis) + (no of multis + 1 (for header)) | or less 1 for header if single row
                    if($count -eq 0){
                    $rows = $count + 2
                    }
                    elseif($count -eq 1){
                    $rows = (8 + $count)
                    }
                    else{
                    $rows = (8 * $count) + ($count + 1)
                    }

                $table = $selection.Tables.add(
                $selection.Range,
                $rows,
                2,
                [Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior,
                [Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitWindow
                )

                $table.style = "Grid Table 4 - Accent 1"
                $table.cell(1,1).range.text = "Item"
                $table.cell(1,2).range.text = "Value"

                $i = 2 # 2 as header is row 1
                $y = 1

                    if($count -eq 0){
                    $table.cell($i,1).range.text = "No programs installed"
                    $table.Cell($i,2).Merge($table.Cell($i,1))
                    }

                    foreach($multi in $multis){

                        if($multi.source -ne 'System'){

                            if($count -gt 1){
                            $table.cell($i,1).Merge($table.cell($i, 2))
                            $table.cell($i,1).range.text = "------ BLOCK $y ------"
                            $i++
                            }

                        $table.cell($i,1).range.text = "Source"
                        $table.cell($i,2).range.text = $multi.source
                        $i++

                        $table.cell($i,1).range.text = "Name"
                        $table.cell($i,2).range.text = $multi.name
                        $i++

                        $table.cell($i,1).range.text = "Version"
                        $table.cell($i,2).range.text = $multi.version
                        $i++

                        $table.cell($i,1).range.text = "Location"
                        $table.cell($i,2).range.text = $multi.location
                        $i++

                        $table.cell($i,1).range.text = "Publisher"
                        $table.cell($i,2).range.text = $multi.publisher
                        $i++

                        $table.cell($i,1).range.text = "Install date"

                            if($multi.install_date -ne ''){
                            $installDate =
                            $multi.install_date.Substring(6,2) + "/" +
                            $multi.install_date.Substring(4,2) + "/" +
                            $multi.install_date.Substring(0,4)
                            }
                            else{
                            $installDate = $multi.install_date
                            }

                        $table.cell($i,2).range.text = $multi.installdate
                        $i++

                        $table.cell($i,1).range.text = "Architecture"
                        $table.cell($i,2).range.text = $multi.architecture
                        $i++

                        $table.cell($i,1).range.text = "Comments"
                        $table.cell($i,2).range.text = $multi.comments
                        $i++

                        $y++

                        }

                    }

                $table.Rows.item(1).Headingformat=-1
                $table.ApplyStyleFirstColumn = $false
                $selection.InsertCaption(-2, ": [Registry] Installed Program - User (Detailed)", $null, 1, $false)

                }

            }

        $selection.EndOf(15) | Out-Null
        $selection.MoveDown() | Out-Null
        $selection.InsertNewPage()
    }
    #endregion ReportInstalledPrograms

    #region ReportWMIOperatingSystem
    # BASIC
    $selection.style = 'Heading 1'
    $selection.TypeText("[WMI] Operating System")
    $selection.TypeParagraph()
    $selection.TypeParagraph()

    $selection.Style = 'Normal'
    $selection.TypeText("This data is collected from the Win32_OperatingSystem WMI class. For more information please go to https://docs.microsoft.com/en-gb/windows/desktop/CIMWin32Prov/win32-operatingsystem")
    $selection.TypeParagraph()
    $selection.TypeParagraph()

        if(($xml.info.win32_operatingsystem).count -eq 0){
        Write-Output "[WARNING] WMI Operating System details not found. Moving on to next section..."
        $warnings ++
        $selection.Style = 'Normal'
        $selection.Font.Color="255"
        $selection.TypeText("[WARNING] WMI Operating System data not found!")
        $selection.TypeParagraph()
        }
        elseif(($xml.info.win32_operatingsystem.errorcode) -eq 1){
        Write-Output "[WARNING] WMI Operating System details not collected. Moving on to next section..."
        $warnings ++
        $selection.Style = 'Normal'
        $selection.Font.Color="255"
        $selection.TypeText("[WARNING] WMI Operating System data not collected!")
        $selection.TypeParagraph()
        $selection.TypeText("Reason for error: $($xml.info.win32_operatingsystem.errortext)")
        $selection.TypeParagraph()
        }
        else{
        Write-Output "[INFO] Populating WMI Operating System table."

        $table = $selection.Tables.add(
        $selection.Range,
        14,
        2,
        [Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior,
        [Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitWindow
        )

        $table.style = "Grid Table 4 - Accent 1"
        $table.cell(1,1).range.text = "Item"
        $table.cell(1,2).range.text = "Value"

        $table.cell(2,1).range.text = "Build Number"
        $table.cell(2,2).range.text = $xml.info.win32_operatingsystem.buildnumber

        $table.cell(3,1).range.text = "Caption"
        $table.cell(3,2).range.text = $xml.info.win32_operatingsystem.caption

        $table.cell(4,1).range.text = "CS Name"
        $table.cell(4,2).range.text = $xml.info.win32_operatingsystem.csname

        $table.cell(5,1).range.text = "Encryption Level"
        $table.cell(5,2).range.text = $xml.info.win32_operatingsystem.encryptionlevel

        $table.cell(6,1).range.text = "Install Date"

        $installDate =
        ($xml.info.win32_operatingsystem.installdate).Substring(6,2) + "/" +
        ($xml.info.win32_operatingsystem.installdate).Substring(4,2) + "/" +
        ($xml.info.win32_operatingsystem.installdate).Substring(0,4) + " " +
        ($xml.info.win32_operatingsystem.installdate).Substring(8,2) + ":" +
        ($xml.info.win32_operatingsystem.installdate).Substring(10,2)

        $table.cell(6,2).range.text = $installDate

        $table.cell(7,1).range.text = "Operating System SKU"

            if($xml.info.win32_operatingsystem.operatingsystemsku -eq 0){
            $sku = "Undefined"
            }
            elseif($xml.info.win32_operatingsystem.operatingsystemsku -eq 1){
            $sku = "Ultimate Edition, e.g. Windows Vista Ultimate."
            }
            elseif($xml.info.win32_operatingsystem.operatingsystemsku -eq 2){
            $sku = "Home Basic Edition"
            }
            elseif($xml.info.win32_operatingsystem.operatingsystemsku -eq 3){
            $sku = "Home Premium Edition"
            }
            elseif($xml.info.win32_operatingsystem.operatingsystemsku -eq 4){
            $sku = "Enterprise Edition"
            }
            elseif($xml.info.win32_operatingsystem.operatingsystemsku -eq 6){
            $sku = "Business Edition"
            }
            elseif($xml.info.win32_operatingsystem.operatingsystemsku -eq 7){
            $sku = "Windows Server Standard Edition (Desktop Experience installation)"
            }
            elseif($xml.info.win32_operatingsystem.operatingsystemsku -eq 8){
            $sku = "Windows Server Datacenter Edition (Desktop Experience installation)"
            }
            elseif($xml.info.win32_operatingsystem.operatingsystemsku -eq 9){
            $sku = "Small Business Server Edition"
            }
            elseif($xml.info.win32_operatingsystem.operatingsystemsku -eq 10){
            $sku = "Enterprise Server Edition"
            }
            elseif($xml.info.win32_operatingsystem.operatingsystemsku -eq 11){
            $sku = "Starter Edition"
            }
            elseif($xml.info.win32_operatingsystem.operatingsystemsku -eq 12){
            $sku = "Datacenter Server Core Edition"
            }
            elseif($xml.info.win32_operatingsystem.operatingsystemsku -eq 13){
            $sku = "Standard Server Core Edition"
            }
            elseif($xml.info.win32_operatingsystem.operatingsystemsku -eq 14){
            $sku = "Enterprise Server Core Edition"
            }
            elseif($xml.info.win32_operatingsystem.operatingsystemsku -eq 17){
            $sku = "Web Server Edition"
            }
            elseif($xml.info.win32_operatingsystem.operatingsystemsku -eq 19){
            $sku = "Home Server Edition"
            }
            elseif($xml.info.win32_operatingsystem.operatingsystemsku -eq 20){
            $sku = "Storage Express Server Edition"
            }
            elseif($xml.info.win32_operatingsystem.operatingsystemsku -eq 21){
            $sku = "Windows Storage Server Standard Edition (Desktop Experience installation)"
            }
            elseif($xml.info.win32_operatingsystem.operatingsystemsku -eq 22){
            $sku = "Windows Storage Server Workgroup Edition (Desktop Experience installation)"
            }
            elseif($xml.info.win32_operatingsystem.operatingsystemsku -eq 23){
            $sku = "Storage Enterprise Server Edition"
            }
            elseif($xml.info.win32_operatingsystem.operatingsystemsku -eq 24){
            $sku = "Server For Small Business Edition"
            }
            elseif($xml.info.win32_operatingsystem.operatingsystemsku -eq 25){
            $sku = "Small Business Server Premium Edition"
            }
            elseif($xml.info.win32_operatingsystem.operatingsystemsku -eq 27){
            $sku = "Windows Enterprise Edition"
            }
            elseif($xml.info.win32_operatingsystem.operatingsystemsku -eq 28){
            $sku = "Windows Ultimate Edition"
            }
            elseif($xml.info.win32_operatingsystem.operatingsystemsku -eq 29){
            $sku = "Windows Server Web Server Edition (Server Core installation)"
            }
            elseif($xml.info.win32_operatingsystem.operatingsystemsku -eq 36){
            $sku = "Windows Server Standard Edition without Hyper-V"
            }
            elseif($xml.info.win32_operatingsystem.operatingsystemsku -eq 37){
            $sku = "Windows Server Datacenter Edition without Hyper-V (full installation)"
            }
            elseif($xml.info.win32_operatingsystem.operatingsystemsku -eq 38){
            $sku = "Windows Server Enterprise Edition without Hyper-V (full installation)"
            }
            elseif($xml.info.win32_operatingsystem.operatingsystemsku -eq 39){
            $sku = "Windows Server Datacenter Edition without Hyper-V (Server Core installation)"
            }
            elseif($xml.info.win32_operatingsystem.operatingsystemsku -eq 40){
            $sku = "Windows Server Standard Edition without Hyper-V (Server Core installation)"
            }
            elseif($xml.info.win32_operatingsystem.operatingsystemsku -eq 41){
            $sku = "Windows Server Enterprise Edition without Hyper-V (Server Core installation)"
            }
            elseif($xml.info.win32_operatingsystem.operatingsystemsku -eq 42){
            $sku = "Microsoft Hyper-V Server"
            }
            elseif($xml.info.win32_operatingsystem.operatingsystemsku -eq 43){
            $sku = "Storage Server Express Edition (Server Core installation)"
            }
            elseif($xml.info.win32_operatingsystem.operatingsystemsku -eq 44){
            $sku = "Storage Server Standard Edition (Server Core installation)"
            }
            elseif($xml.info.win32_operatingsystem.operatingsystemsku -eq 450){
            $sku = "Storage Server Workgroup Edition (Server Core installation)"
            }
            elseif($xml.info.win32_operatingsystem.operatingsystemsku -eq 46){
            $sku = "Storage Server Enterprise Edition (Server Core installation)"
            }
            elseif($xml.info.win32_operatingsystem.operatingsystemsku -eq 50){
            $sku = "Windows Server Essentials (Desktop Experience installation)"
            }
            elseif($xml.info.win32_operatingsystem.operatingsystemsku -eq 63){
            $sku = "Small Business Server Premium (Server Core installation)"
            }
            elseif($xml.info.win32_operatingsystem.operatingsystemsku -eq 64){
            $sku = "Windows Compute Cluster Server without Hyper-V"
            }
            elseif($xml.info.win32_operatingsystem.operatingsystemsku -eq 97){
            $sku = "Windows RT"
            }
            elseif($xml.info.win32_operatingsystem.operatingsystemsku -eq 101){
            $sku = "Windows Home"
            }
            elseif($xml.info.win32_operatingsystem.operatingsystemsku -eq 103){
            $sku = "Windows Professional with Media Center"
            }
            elseif($xml.info.win32_operatingsystem.operatingsystemsku -eq 104){
            $sku = "Windows Mobile"
            }
            elseif($xml.info.win32_operatingsystem.operatingsystemsku -eq 123){
            $sku = "Windows IoT (Internet of Things) Core"
            }
            elseif($xml.info.win32_operatingsystem.operatingsystemsku -eq 143){
            $sku = "Windows Server Datacenter Edition (Nano Server installation)"
            }
            elseif($xml.info.win32_operatingsystem.operatingsystemsku -eq 144){
            $sku = "Windows Server Standard Edition (Nano Server installation)"
            }
            elseif($xml.info.win32_operatingsystem.operatingsystemsku -eq 147){
            $sku = "Windows Server Datacenter Edition (Server Core installation)"
            }
            elseif($xml.info.win32_operatingsystem.operatingsystemsku -eq 148){
            $sku = "Windows Server Standard Edition (Server Core installation)"
            }
            else{
            $sku = "$sku [UNKNOWN]"
            }

        $table.cell(7,2).range.text = $sku

        $table.cell(8,1).range.text = "OS Architecture"
        $table.cell(8,2).range.text = $xml.info.win32_operatingsystem.osarchitecture

        $table.cell(9,1).range.text = "OS Product Suite"

        # Convert to bits and pad to 16
        $bits = ([Int][Convert]::ToString($xml.info.win32_operatingsystem.osproductsuite,2)).ToString("0000000000000000")

        # Convert to char array
        $chars = $bits.ToCharArray()
        # Create blank array
        $mappings = @()

            # If bit is set, map value to $mapping. 16 bit binary number (left to right)
            # https://docs.microsoft.com/en-gb/windows/desktop/WmiSdk/bitmap-and-bitvalues
            # Decimal 32768, Binary 1000 0000 0000 0000â€¬
            if($chars[0] -eq "1"){
            $mapping = New-Object System.Object
            $mapping | Add-Member -type NoteProperty -name Name -Value "- [NOT USED].`r`n"
            $mappings += $mapping
            }
            # Decimal 16384, Binary 0100 0000 0000 0000
            if($chars[1] -eq "1"){
            $mapping = New-Object System.Object
            $mapping | Add-Member -type NoteProperty -name Name -Value "- Compute Cluster Edition is installed.`r`n"
            $mappings += $mapping
            }
            # Decimal 8192, Binary 0010 0000 0000 0000
            if($chars[2] -eq "1"){
            $mapping = New-Object System.Object
            $mapping | Add-Member -type NoteProperty -name Name -Value "- Storage Server Edition is installed.`r`n"
            $mappings += $mapping
            }
            # Decimal 4096, Binary 0001 0000 0000 0000
            if($chars[3] -eq "1"){
            $mapping = New-Object System.Object
            $mapping | Add-Member -type NoteProperty -name Name -Value "- [NOT USED].`r`n"
            $mappings += $mapping
            }
            # Decimal 2048, Binary 0000 1000 0000 0000
            if($chars[4] -eq "1"){
            $mapping = New-Object System.Object
            $mapping | Add-Member -type NoteProperty -name Name -Value "- [NOT USED].`r`n"
            $mappings += $mapping
            }
            # Decimal 1024, Binary 0000 0100 0000 0000
            if($chars[5] -eq "1"){
            $mapping = New-Object System.Object
            $mapping | Add-Member -type NoteProperty -name Name -Value "- Web Server Edition is installed.`r`n"
            $mappings += $mapping
            }
            # Decimal 512, Binary 0000 0010 0000 0000
            if($chars[6] -eq "1"){
            $mapping = New-Object System.Object
            $mapping | Add-Member -type NoteProperty -name Name -Value "- Windows Home Edition is installed.`r`n"
            $mappings += $mapping
            }
            # Decimal 256, Binary 0000 0001 0000 0000
            if($chars[7] -eq "1"){
            $mapping = New-Object System.Object
            $mapping | Add-Member -type NoteProperty -name Name -Value "- Terminal Services is installed, but only one interactive session is supported.`r`n"
            $mappings += $mapping
            }
            # Decimal 128, Binary 0000 0000 1000 0000
            if($chars[8] -eq "1"){
            $mapping = New-Object System.Object
            $mapping | Add-Member -type NoteProperty -name Name -Value "- A Datacenter edition is installed.`r`n"
            $mappings += $mapping
            }
            # Decimal 64, Binary 0000 0000 0100 0000
            if($chars[9] -eq "1"){
            $mapping = New-Object System.Object
            $mapping | Add-Member -type NoteProperty -name Name -Value "- Windows Embedded is installed.`r`n"
            $mappings += $mapping
            }
            # Decimal 32, Binary 0000 0000 0010 0000
            if($chars[10] -eq "1"){
            $mapping = New-Object System.Object
            $mapping | Add-Member -type NoteProperty -name Name -Value "- Microsoft Small Business Server is installed with the restrictive client license.`r`n"
            $mappings += $mapping
            }
            # Decimal 16, Binary 0000 0000 0001 0000
            if($chars[11] -eq "1"){
            $mapping = New-Object System.Object
            $mapping | Add-Member -type NoteProperty -name Name -Value "- Terminal Services is installed.`r`n"
            $mappings += $mapping
            }
            # Decimal 8, Binary 0000 0000 0000 1000
            if($chars[12] -eq "1"){
            $mapping = New-Object System.Object
            $mapping | Add-Member -type NoteProperty -name Name -Value "- Communication Server is installed.`r`n"
            $mappings += $mapping
            }
            # Decimal 4, Binary 0000 0000 0000 0100
            if($chars[13] -eq "1"){
            $mapping = New-Object System.Object
            $mapping | Add-Member -type NoteProperty -name Name -Value "- Windows BackOffice components are installed.`r`n"
            $mappings += $mapping
            }
            # Decimal 2, Binary 0000 0000 0000 0010
            if($chars[14] -eq "1"){
            $mapping = New-Object System.Object
            $mapping | Add-Member -type NoteProperty -name Name -Value "- Windows Server 2008 Enterprise is installed.`r`n"
            $mappings += $mapping
            }
            # Decimal 1, Binary 0000 0000 0000 0001
            if($chars[15] -eq "1"){
            $mapping = New-Object System.Object
            $mapping | Add-Member -type NoteProperty -name Name -Value "- Microsoft Small Business Server was once installed, but may have been upgraded to another version of Windows."
            $mappings += $mapping
            }

        $a = $mappings.name

        $table.cell(9,2).range.text = "$a"

        $table.cell(10,1).range.text = "Product Type"

            if($xml.info.win32_operatingsystem.producttype -eq 1){
            $productType = "Workstation"
            }
            elseif($xml.info.win32_operatingsystem.producttype -eq 2){
            $productType = "Domain Controller"
            }
            elseif($xml.info.win32_operatingsystem.producttype -eq 3){
            $productType = "Server"
            }
            else{
            $productType = "$productType [UNKNOWN]"
            }

        $table.cell(10,2).range.text = $productType
        $table.cell(11,1).range.text = "Service Pack Major Version"
        $table.cell(11,2).range.text = $xml.info.win32_operatingsystem.servicepackmajorversion
        $table.cell(12,1).range.text = "Service Pack Minor Version"
        $table.cell(12,2).range.text = $xml.info.win32_operatingsystem.servicepackminorversion
        $table.cell(13,1).range.text = "Total Visible Memory Size (GB)"

        # Kilobtyes to Gigabytes
        $a = [system.decimal](($xml.info.win32_operatingsystem.totalvisiblememorysize)/1000000)
        $b = [math]::Round($a, 2)

        $table.cell(13,2).range.text = "$b"
        $table.cell(14,1).range.text = "Version"
        $table.cell(14,2).range.text = $xml.info.win32_operatingsystem.version

        $table.Rows.item(1).Headingformat=-1
        $table.ApplyStyleFirstColumn = $false
        $selection.InsertCaption(-2, ": [WMI] Operating System", $null, 1, $false)

        }

    $selection.EndOf(15) | Out-Null
    $selection.MoveDown() | Out-Null
    $selection.InsertNewPage()
    #endregion ReportWMIOperatingSystem

    #region ReportWMIComputerSystem
    # BASIC
    $selection.style = 'Heading 1'
    $selection.TypeText("[WMI] Computer System")
    $selection.TypeParagraph()
    $selection.TypeParagraph()

    $selection.Style = 'Normal'
    $selection.TypeText("This data is collected from the Win32_ComputerSystem WMI class. For more information please go to https://docs.microsoft.com/en-gb/windows/desktop/CIMWin32Prov/win32-computersystem")
    $selection.TypeParagraph()
    $selection.TypeParagraph()

        if(($xml.info.win32_computersystem).count -eq 0){
        Write-Output "[WARNING] WMI Computer System details not found. Moving on to next section..."
        $warnings ++
        $selection.Style = 'Normal'
        $selection.Font.Color="255"
        $selection.TypeText("[WARNING] WMI Computer System data not found!")
        $selection.TypeParagraph()
        }
        elseif(($xml.info.win32_computersystem.errorcode) -eq 1){
        Write-Output "[WARNING] WMI Computer System details not collected. Moving on to next section..."
        $warnings ++
        $selection.Style = 'Normal'
        $selection.Font.Color="255"
        $selection.TypeText("[WARNING] WMI Computer System data not collected!")
        $selection.TypeParagraph()
        $selection.TypeText("Reason for error: $($xml.info.win32_computersystem.errortext)")
        $selection.TypeParagraph()
        }
        else{
        Write-Output "[INFO] Populating WMI Computer System table."

        $table = $selection.Tables.add(
        $selection.Range,
        13,
        2,
        [Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior,
        [Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitWindow
        )

        $table.style = "Grid Table 4 - Accent 1"
        $table.cell(1,1).range.text = "Item"
        $table.cell(1,2).range.text = "Value"

        $table.cell(2,1).range.text = "Admin Password Status"

            if($xml.info.win32_computersystem.adminpasswordstatus -eq 0){
            $admPwStatus = "Disabled"
            }
            elseif($xml.info.win32_computersystem.adminpasswordstatus -eq 1){
            $admPwStatus = "Enabled"
            }
            elseif($xml.info.win32_computersystem.adminpasswordstatus -eq 2){
            $admPwStatus = "Not Implemented"
            }
            elseif($xml.info.win32_computersystem.adminpasswordstatus -eq 3){
            $admPwStatus = "Unknown"
            }
            else{
            $admPwStatus = "$admPwStatus [UNKNOWN]"
            }

        $table.cell(2,2).range.text = $admPwStatus

        $table.cell(3,1).range.text = "Name"
        $table.cell(3,2).range.text = $xml.info.win32_computersystem.name

        $table.cell(4,1).range.text = "Domain"
        $table.cell(4,2).range.text = $xml.info.win32_computersystem.domain

        $table.cell(5,1).range.text = "Domain Role"

            if($xml.info.win32_computersystem.domainrole -eq 0){
            $domainRole = "Standalone Workstation"
            }
            elseif($xml.info.win32_computersystem.domainrole -eq 1){
            $domainRole = "Member Workstation"
            }
            elseif($xml.info.win32_computersystem.domainrole -eq 2){
            $domainRole = "Standalone Server"
            }
            elseif($xml.info.win32_computersystem.domainrole -eq 3){
            $domainRole = "Member Server"
            }
            elseif($xml.info.win32_computersystem.domainrole -eq 4){
            $domainRole = "Backup Domain Controller"
            }
            elseif($xml.info.win32_computersystem.domainrole -eq 5){
            $domainRole = "Primary Domain Controller"
            }
            else{
            $domainRole = "$domainRole [UNKNOWN]"
            }

        $table.cell(5,2).range.text = $domainRole

        $table.cell(6,1).range.text = "Manufacturer"
        $table.cell(6,2).range.text = $xml.info.win32_computersystem.manufacturer

        $table.cell(7,1).range.text = "Model"
        $table.cell(7,2).range.text = $xml.info.win32_computersystem.model

        $table.cell(8,1).range.text = "Number Of Logical Processors"
        $table.cell(8,2).range.text = $xml.info.win32_computersystem.numberoflogicalprocessors

        $table.cell(9,1).range.text = "Part Of Domain"
        $table.cell(9,2).range.text = $xml.info.win32_computersystem.partofdomain

        $table.cell(10,1).range.text = "Roles"

        $array = @($xml.info.win32_computersystem.roles.split(" "))
        $roles = @()

            foreach ($a in $array){

            $role = New-Object System.Object

            $role | Add-Member -type NoteProperty -name v -Value "$a`r`n"
            $roles += $role

            }

        # entries (except 1st) start with a space, remove with -replace
        $r = [string]($roles.v) -replace " ", ""

        $table.cell(10,2).range.text = "$r"

        $table.cell(11,1).range.text = "System Type"
        $table.cell(11,2).range.text = $xml.info.win32_computersystem.systemtype

        $table.cell(12,1).range.text = "Total Physical Memory (GB)"

        # Convert bytes to GB and Round
        [string]$a = [system.decimal](($xml.info.win32_computersystem.totalphysicalmemory)/1073741824)
        $b = [math]::Round($a)

        $table.cell(12,2).range.text = "$b" # ""'s needed otherwise it throws an error.

        $table.cell(13,1).range.text = "Wake Up Type"

            if($xml.info.win32_computersystem.wakeuptype -eq 0){
            $wakeUpType = "Reserved"
            }
            elseif($xml.info.win32_computersystem.wakeuptype -eq 1){
            $wakeUpType = "Other"
            }
            elseif($xml.info.win32_computersystem.wakeuptype -eq 2){
            $wakeUpType = "Unknown"
            }
            elseif($xml.info.win32_computersystem.wakeuptype -eq 3){
            $wakeUpType = "APM Timer"
            }
            elseif($xml.info.win32_computersystem.wakeuptype -eq 4){
            $wakeUpType = "Modem Ring"
            }
            elseif($xml.info.win32_computersystem.wakeuptype -eq 5){
            $wakeUpType = "LAN Remote"
            }
	        elseif($xml.info.win32_computersystem.wakeuptype -eq 6){
            $wakeUpType = "Power Switch"
            }
	        elseif($xml.info.win32_computersystem.wakeuptype -eq 7){
            $wakeUpType = "PCI PME#" # is part of the name
            }
	        elseif($xml.info.win32_computersystem.wakeuptype -eq 8){
            $wakeUpType = "AC Power Restored"
            }
            else{
            $wakeUpType = "$wakeUpType [UNKNOWN]"
            }

        $table.cell(13,2).range.text = $wakeUpType

        $table.Rows.item(1).Headingformat=-1
        $table.ApplyStyleFirstColumn = $false
        $selection.InsertCaption(-2, ": [WMI] Computer System", $null, 1, $false)

        }

    $selection.EndOf(15) | Out-Null
    $selection.MoveDown() | Out-Null
    $selection.InsertNewPage()
    #endregion ReportWMIComputerSystem

    #region ReportWMIWinSat
    # BASIC
    $selection.style = 'Heading 1'
    $selection.TypeText("[WMI] Windows System Assessment Tool")
    $selection.TypeParagraph()
    $selection.TypeParagraph()

    $selection.Style = 'Normal'
    $selection.TypeText("This data is collected from the Win32_WinSat WMI class. For more information please go to https://docs.microsoft.com/en-us/windows/win32/winsat/win32-winsat")
    $selection.TypeParagraph()
    $selection.TypeParagraph()

        if(($xml.info.win32_winsat).count -eq 0){
        Write-Output "[WARNING] WMI WinSat details not found. Moving on to next section..."
        $warnings ++
        $selection.Style = 'Normal'
        $selection.Font.Color="255"
        $selection.TypeText("[WARNING] WMI WinSat data not found!")
        $selection.TypeParagraph()
        }
        elseif(($xml.info.win32_winsat.errorcode) -eq 1){
        Write-Output "[WARNING] WMI WinSat details not collected. Moving on to next section..."
        $warnings ++
        $selection.Style = 'Normal'
        $selection.Font.Color="255"
        $selection.TypeText("[WARNING] WMI WinSat data not collected!")
        $selection.TypeParagraph()
        $selection.TypeText("Reason for error: $($xml.info.win32_winsat.errortext)")
        $selection.TypeParagraph()
        }
        else{
        Write-Output "[INFO] Populating WMI WinSat table."

        $table = $selection.Tables.add(
        $selection.Range,
        9,
        2,
        [Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior,
        [Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitWindow
        )

        $table.style = "Grid Table 4 - Accent 1"
        $table.cell(1,1).range.text = "Item"
        $table.cell(1,2).range.text = "Value"

        $table.cell(2,1).range.text = "CPU Score"
        $table.cell(2,2).range.text = $xml.info.win32_winsat.cpuscore

        $table.cell(3,1).range.text = "D3D Score"
        $table.cell(3,2).range.text = $xml.info.win32_winsat.d3dscore

        $table.cell(4,1).range.text = "Disk Score"
        $table.cell(4,2).range.text = $xml.info.win32_winsat.diskscore

        $table.cell(5,1).range.text = "Graphics Score"
        $table.cell(5,2).range.text = $xml.info.win32_winsat.graphicsscore

        $table.cell(6,1).range.text = "Memory Score"
        $table.cell(6,2).range.text = $xml.info.win32_winsat.memoryscore

        $table.cell(7,1).range.text = "Time Taken"
        $table.cell(7,2).range.text = $xml.info.win32_winsat.timetaken

        $table.cell(8,1).range.text = "WinSat Assessment State"

            if($xml.info.win32_winsat.winsatassessmentstate -eq 0){
            $winSatAssessmentState = "State Unknown"
            }
            elseif($xml.info.win32_winsat.winsatassessmentstate -eq 1){
            $winSatAssessmentState = "Valid"
            }
            elseif($xml.info.win32_winsat.winsatassessmentstate -eq 2){
            $winSatAssessmentState = "Incoherent With Hardware"
            }
            elseif($xml.info.win32_winsat.winsatassessmentstate -eq 3){
            $winSatAssessmentState = "No Assessement Available"
            }
            elseif($xml.info.win32_winsat.winsatassessmentstate -eq 4){
            $winSatAssessmentState = "Invalid"
            }
            else{
            $winSatAssessmentState = "$winSatAssessmentState [UNKNOWN]"
            }

        $table.cell(8,2).range.text = "$winSatAssessmentState"

        $table.cell(9,1).range.text = "Win SPR Level"
        $table.cell(9,2).range.text = $xml.info.win32_winsat.winsprlevel

    $table.Rows.item(1).Headingformat=-1
    $table.ApplyStyleFirstColumn = $false
    $selection.InsertCaption(-2, ": [WMI] Windows System Assessment Tool", $null, 1, $false)

    }

    $selection.EndOf(15) | Out-Null
    $selection.MoveDown() | Out-Null
    $selection.InsertNewPage()
    #endregion ReportWMIWinSat

    #region ReportWMINetworkAdapterConfiguration
    # BASIC
    $selection.style = 'Heading 1'
    $selection.TypeText("[WMI] Network Adapter Configuration")
    $selection.TypeParagraph()
    $selection.TypeParagraph()

    $selection.Style = 'Normal'
    $selection.TypeText("This data is collected from the Win32_NetworkAdapterConfiguration WMI class. For more information please go to https://docs.microsoft.com/en-gb/windows/desktop/CIMWin32Prov/win32-networkadapterconfiguration")
    $selection.TypeParagraph()
    $selection.Style = 'Normal'
    $selection.TypeText("Note: There is a known bug in this method for collecting DHCP lease times for this Operating System. Times below marked between *** *** need to be validated using other means. For more information please go to https://social.technet.microsoft.com/Forums/windows/en-US/e22b8272-aafa-484b-a68f-1347c71381b0/wmi-cim-dhcp-information-wrong?forum=win10itprogeneral")
    $selection.TypeParagraph()
    $selection.TypeParagraph()

        if(($xml.info.win32_networkadapterconfiguration).count -eq 0){
        Write-Output "[WARNING] WMI Network Adapter Configuration details not found. Moving on to next section..."
        $warnings ++
        $selection.Style = 'Normal'
        $selection.Font.Color="255"
        $selection.TypeText("[WARNING] WMI Network Adapter Configuration data not found!")
        $selection.TypeParagraph()
        }
        elseif(($xml.info.win32_networkadapterconfiguration.errorcode) -eq 1){
        Write-Output "[WARNING] WMI Network Adapter Configuration details not collected. Moving on to next section..."
        $warnings ++
        $selection.Style = 'Normal'
        $selection.Font.Color="255"
        $selection.TypeText("[WARNING] WMI Network Adapter Configuration data not collected!")
        $selection.TypeParagraph()
        $selection.TypeText("Reason for error: $($xml.info.win32_networkadapterconfiguration.errortext)")
        $selection.TypeParagraph()
        }
        else{
        Write-Output "[INFO] Populating WMI Network Adapater Configuration table."

        # Count rows in multi
        $multis = $xml.selectnodes("//info/win32_networkadapterconfiguration/win32_networkadapterconfiguration_multi")
        $count = ($multis | Measure-Object).Count

            # Calculate rows. (No of items x no of multis) + (no of multis + 1 (for header)) | or less 1 for header if single row
            if($count -eq 1){
            $rows = (20 + $count)
            }
            else{
            $rows = (20 * $count) + ($count + 1)
            }

        $table = $selection.Tables.add(
        $selection.Range,
        $rows,
        2,
        [Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior,
        [Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitWindow
        )

        $table.style = "Grid Table 4 - Accent 1"
        $table.cell(1,1).range.text = "Item"
        $table.cell(1,2).range.text = "Value"

        $i = 2 # 2 as header is row 1
        $y = 1

            foreach($multi in $multis){

                if($count -gt 1){
                $table.cell($i,1).Merge($table.cell($i, 2))
                $table.cell($i,1).range.text = "------ BLOCK $y ------"
                $i++
                }

                # If blank dont format using substring or an error will be thrown
                if($multi.dhcpleaseexpires -eq ""){
                $dHCPLE = ""
                }
                else{
                $dHCPLE =
                ($multi.dhcpleaseexpires).Substring(6,2) + "/" +
                ($multi.dhcpleaseexpires).Substring(4,2) + "/" +
                ($multi.dhcpleaseexpires).Substring(0,4) + " " +
                ($multi.dhcpleaseexpires).Substring(8,2) + ":" +
                ($multi.dhcpleaseexpires).Substring(10,2)
                }

            $table.cell($i,1).range.text = "DHCP Lease Expires"
            $table.cell($i,2).range.font.textColor="255"
            $table.cell($i,2).range.text = " *** $dHCPLE ***"

            $i++

            $table.cell($i,1).range.text = "Description"
            $table.cell($i,2).range.text = $multi.description
            $i++

            $table.cell($i,1).range.text = "DHCP Enabled"
            $table.cell($i,2).range.text = $multi.dhcpenabled
            $i++

                # If blank dont format using substring or an error will be thrown
                if($multi.dhcpleaseobtained -eq ""){
                $dHCPLO = ""
                }
                else{
                $dHCPLO =
                ($multi.dhcpleaseobtained).Substring(6,2) + "/" +
                ($multi.dhcpleaseobtained).Substring(4,2) + "/" +
                ($multi.dhcpleaseobtained).Substring(0,4) + " " +
                ($multi.dhcpleaseobtained).Substring(8,2) + ":" +
                ($multi.dhcpleaseobtained).Substring(10,2)
                }

            $table.cell($i,1).range.text = "DHCP Lease Obtained"
            $table.cell($i,2).range.font.textColor="255"
            $table.cell($i,2).range.text = " *** $dHCPLO ***"

            $i++

            $table.cell($i,1).range.text = "DHCP Server"
            $table.cell($i,2).range.text = $multi.dhcpserver
            $i++

            $table.cell($i,1).range.text = "DNS Domain"
            $table.cell($i,2).range.text = $multi.dnsdomain
            $i++

            $table.cell($i,1).range.text = "DNS Domain Suffix Search Order"
            $table.cell($i,2).range.text = $multi.dnsdomainsuffixsearchorder -replace " ", "`r`n"
            $i++

            $table.cell($i,1).range.text = "DNS Enabled For WINS Resolution"
            $table.cell($i,2).range.text = $multi.dnsenabledforwinsresolution
            $i++

            $table.cell($i,1).range.text = "DNS Hostname"
            $table.cell($i,2).range.text = $multi.dnshostname
            $i++

            $table.cell($i,1).range.text = "DNS Server Search Order"
            $table.cell($i,2).range.text = $multi.dnsserversearchorder -replace " ", "`r`n"
            $i++

            $table.cell($i,1).range.text = "IP Address"
            $table.cell($i,2).range.text = $multi.ipaddress -replace " ", "`r`n"
            $i++

            $table.cell($i,1).range.text = "IP Enabled"
            $table.cell($i,2).range.text = $multi.ipenabled
            $i++

            $table.cell($i,1).range.text = "IP Filter Security Enabled"
            $table.cell($i,2).range.text = $multi.ipfiltersecurityenabled
            $i++

            $table.cell($i,1).range.text = "WINS Enable LMhosts Lookup"
            $table.cell($i,2).range.text = $multi.winsenablelmhostslookup
            $i++

            $table.cell($i,1).range.text = "WINS Primary Server"
            $table.cell($i,2).range.text = $multi.winsprimaryserver
            $i++

            $table.cell($i,1).range.text = "WINS Secondary Server"
            $table.cell($i,2).range.text = $multi.winssecondaryserver
            $i++

            $table.cell($i,1).range.text = "Caption"
            $table.cell($i,2).range.text = $multi.caption
            $i++

            $table.cell($i,1).range.text = "Default IP Gateway"
            $table.cell($i,2).range.text = $multi.defaultipgateway
            $i++

            $table.cell($i,1).range.text = "IP Subnet"
            $table.cell($i,2).range.text = $multi.ipsubnet
            $i++

            $table.cell($i,1).range.text = "MAC Address"
            $table.cell($i,2).range.text = $multi.macaddress
            $i++

            $y++

            }

        $table.Rows.item(1).Headingformat=-1
        $table.ApplyStyleFirstColumn = $false
        $selection.InsertCaption(-2, ": [WMI] Network Adapter Configuration", $null, 1, $false)

        }

    $selection.EndOf(15) | Out-Null
    $selection.MoveDown() | Out-Null
    $selection.InsertNewPage()
    #endregion ReportWMINetworkAdapterConfiguration

    #region ReportWMIPhysicalMemory
    #BASIC & DETAILED
    $selection.style = 'Heading 1'
    $selection.TypeText("[WMI] Physical Memory")
    $selection.TypeParagraph()
    $selection.TypeParagraph()

    $selection.Style = 'Normal'
    $selection.TypeText("This data is collected from the Win32_PhysicalMemory WMI class. For more information please go to https://docs.microsoft.com/en-gb/windows/desktop/CIMWin32Prov/win32-physicalmemory")
    $selection.TypeParagraph()
    $selection.TypeParagraph()

        if($reportType -ne 'Basic'){
        $selection.style = 'Heading 2'
        $selection.TypeText("Basic")
        $selection.TypeParagraph()
        $selection.TypeParagraph()
        }

        if(($xml.info.win32_physicalmemory).count -eq 0){
        Write-Output "[WARNING] WMI Physical Memory (Basic) details not found. Moving on to next section..."
        $warnings ++
        $selection.Style = 'Normal'
        $selection.Font.Color="255"
        $selection.TypeText("[WARNING] WMI Physical Memory (Basic) data not found!")
        $selection.TypeParagraph()
        }
        elseif(($xml.info.win32_physicalmemory.errorcode) -eq 1){
        Write-Output "[WARNING] WMI Physical Memory details not collected. Moving on to next section..."
        $warnings ++
        $selection.Style = 'Normal'
        $selection.Font.Color="255"
        $selection.TypeText("[WARNING] WMI Physical Memory data not collected!")
        $selection.TypeParagraph()
        $selection.TypeText("Reason for error: $($xml.info.win32_physicalmemory.errortext)")
        $selection.TypeParagraph()
        }
        else{
        Write-Output "[INFO] Populating WMI Physical Memory (Basic) table."

        # Count rows in multi
        $multis = $xml.selectnodes("//info/win32_physicalmemory/win32_physicalmemory_multi")
        $count = ($multis | Measure-Object).Count

        # Calculate rows (from query above) and add 1 for the header
        $rows = $count + 1

        $table = $selection.Tables.add(
        $selection.Range,
        $rows,
        4,
        [Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior,
        [Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitWindow
        )

        $table.style = "Grid Table 4 - Accent 1"
        $table.cell(1,1).range.text = "Device Locator"
        $table.cell(1,2).range.text = "Memory Type"
        $table.cell(1,3).range.text = "Configured Clock Speed (MHz)"
        $table.cell(1,4).range.text = "Capacity (GB)"

        $i = 2 # 2 as header is row 1

            foreach($multi in $multis){

            $table.cell($i,1).range.text = $multi.devicelocator

                if($multi.memorytype -eq 0){
                $memoryType = "Unknown"
                }
                elseif($multi.memorytype -eq 1){
                $memoryType = "Other"
                }
                elseif($multi.memorytype -eq 2){
                $memoryType = "DRAM"
                }
                elseif($multi.memorytype -eq 3){
                $memoryType = "Synchronous DRAM"
                }
                elseif($multi.memorytype -eq 4){
                $memoryType = "Cache DRAM"
                }
                elseif($multi.memorytype -eq 5){
                $memoryType = "EDO"
                }
	            elseif($multi.memorytype -eq 6){
                $memoryType = "EDRAM"
                }
	            elseif($multi.memorytype -eq 7){
                $memoryType = "VRAM"
                }
	            elseif($multi.memorytype -eq 8){
                $memoryType = "SRAM"
                }
	            elseif($multi.memorytype -eq 9){
                $memoryType = "RAM"
                }
	            elseif($multi.memorytype -eq 10){
                $memoryType = "ROM"
                }
	            elseif($multi.memorytype -eq 11){
                $memoryType = "Flash"
                }
	            elseif($multi.memorytype -eq 12){
                $memoryType = "EEPROM"
                }
	            elseif($multi.memorytype -eq 13){
                $memoryType = "FEPROM"
                }
	            elseif($multi.memorytype -eq 14){
                $memoryType = "EPROM"
                }
	            elseif($multi.memorytype -eq 15){
                $memoryType = "CDRAM"
                }
	            elseif($multi.memorytype -eq 16){
                $memoryType = "3DRAM"
                }
	            elseif($multi.memorytype -eq 17){
                $memoryType = "SDRAM"
                }
	            elseif($multi.memorytype -eq 18){
                $memoryType = "SGRAM"
                }
	            elseif($multi.memorytype -eq 19){
                $memoryType = "RDRAM"
                }
	            elseif($multi.memorytype -eq 20){
                $memoryType = "DDR"
                }
	            elseif($multi.memorytype -eq 21){
                $memoryType = "DDR2"
                }
		        elseif($multi.memorytype -eq 22){
                $memoryType = "DDR2 FB-DIMM"
                }
		        # No 23
		        elseif($multi.memorytype -eq 24){
                $memoryType = "DDR3"
                }
		        elseif($multi.memorytype -eq 25){
                $memoryType = "FBD2"
                }
		        else{
                $memoryType = "$memoryType [UNKNOWN]"
                }

            $table.cell($i,2).range.text = $memoryType
            $table.cell($i,3).range.text = $multi.configuredclockspeed

            # Convert bytes to GB and Round
            [string]$a = [system.decimal](($multi.capacity)/1073741824)
            $b = [math]::Round($a, 2)

            $table.cell($i,4).range.text = "$b"

            $i++

            }

        $table.Rows.item(1).Headingformat=-1
        $table.ApplyStyleFirstColumn = $false
        $selection.InsertCaption(-2, ": [WMI] Physical Memory (Basic)", $null, 1, $false)

          if($reportType -ne "Basic"){
          $selection.TypeParagraph()
          }

        }

    $selection.EndOf(15) | Out-Null
    $selection.MoveDown() | Out-Null

        if($reportType -eq "Detailed"){
        $selection.TypeParagraph()
        $selection.style = 'Heading 2'
        $selection.TypeText("Detailed")
        $selection.TypeParagraph()
        $selection.TypeParagraph()

            if(($xml.info.win32_physicalmemory).count -eq 0){
            Write-Output "[WARNING] WMI Physical Memory (Detailed) details not found. Moving on to next section..."
            $warnings ++
            $selection.Style = 'Normal'
            $selection.Font.Color="255"
            $selection.TypeText("[WARNING] WMI Physical Memory (Detailed) data not found!")
            $selection.TypeParagraph()
            }
            elseif(($xml.info.win32_physicalmemory.errorcode) -eq 1){
            Write-Output "[WARNING] WMI Physical Memory details not collected. Moving on to next section..."
            $warnings ++
            $selection.Style = 'Normal'
            $selection.Font.Color="255"
            $selection.TypeText("[WARNING] WMI Physical Memory data not collected!")
            $selection.TypeParagraph()
            $selection.TypeText("Reason for error: $($xml.info.win32_physicalmemory.errortext)")
            $selection.TypeParagraph()
            }
            else{
            Write-Output "[INFO] Populating WMI Physical Memory (Detailed) table."

            # Count rows in multi
            $multis = $xml.selectnodes("//info/win32_physicalmemory/win32_physicalmemory_multi")
            $count = ($multis | Measure-Object).Count

                # Calculate rows. (No of items x no of multis) + (no of multis + 1 (for header)) | or less 1 for header if single row
                if($count -eq 1){
                $rows = (11 + $count)
                }
                else{
                $rows = (11 * $count) + ($count + 1)
                }

            $table = $selection.Tables.add(
            $selection.Range,
            $rows,
            2,
            [Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior,
            [Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitWindow
            )

            $table.style = "Grid Table 4 - Accent 1"
            $table.cell(1,1).range.text = "Item"
            $table.cell(1,2).range.text = "Value"

            $i = 2 # 2 as header is row 1
            $y = 1

                foreach($multi in $multis){

                    if($count -gt 1){
                    $table.cell($i,1).Merge($table.cell($i, 2))
                    $table.cell($i,1).range.text = "------ BLOCK $y ------"
                    $i++
                    }

                # Convert bytes to GB and Round
                [string]$a = [system.decimal](($multi.capacity)/1073741824)
                $b = [math]::Round($a, 2)

                $table.cell($i,1).range.text = "Capacity (GB)"
                $table.cell($i,2).range.text = "$b"
                $i++

                $table.cell($i,1).range.text = "Configured Clock Speed (MHz)"
                $table.cell($i,2).range.text = $multi.configuredclockspeed
                $i++

                $table.cell($i,1).range.text = "Data Width (Bits)"
                $table.cell($i,2).range.text = $multi.datawidth
                $i++

                $table.cell($i,1).range.text = "Device Locator"
                $table.cell($i,2).range.text = $multi.devicelocator
                $i++

                    if($multi.formfactor -eq 0){
                    $formFactor = "Unknown"
                    }
                    elseif($multi.formfactor -eq 1){
                    $formFactor = "Other"
                    }
                    elseif($multi.formfactor -eq 2){
                    $formFactor = "SIP"
                    }
                    elseif($multi.formfactor -eq 3){
                    $formFactor = "DIP"
                    }
                    elseif($multi.formfactor -eq 4){
                    $formFactor = "ZIP"
                    }
                    elseif($multi.formfactor -eq 5){
                    $formFactor = "SOJ"
                    }
	                elseif($multi.formfactor -eq 6){
                    $formFactor = "Proprietary"
                    }
	                elseif($multi.formfactor -eq 7){
                    $formFactor = "SIMM"
                    }
	                elseif($multi.formfactor -eq 8){
                    $formFactor = "DIMM"
                    }
	                elseif($multi.formfactor -eq 9){
                    $formFactor = "TSOP"
                    }
	                elseif($multi.formfactor -eq 10){
                    $formFactor = "PGA"
                    }
	                elseif($multi.formfactor -eq 11){
                    $formFactor = "RIMM"
                    }
	                elseif($multi.formfactor -eq 12){
                    $formFactor = "SODIMM"
                    }
	                elseif($multi.formfactor -eq 13){
                    $formFactor = "SRIMM"
                    }
	                elseif($multi.formfactor -eq 14){
                    $formFactor = "SMD"
                    }
	                elseif($multi.formfactor -eq 15){
                    $formFactor = "SSMP"
                    }
	                elseif($multi.formfactor -eq 16){
                    $formFactor = "QFP"
                    }
	                elseif($multi.formfactor -eq 17){
                    $formFactor = "TQFP"
                    }
	                elseif($multi.formfactor -eq 18){
                    $formFactor = "SOIC"
                    }
	                elseif($multi.formfactor -eq 19){
                    $formFactor = "LCC"
                    }
	                elseif($multi.formfactor -eq 20){
                    $formFactor = "PLCC"
                    }
	                elseif($multi.formfactor -eq 21){
                    $formFactor = "BGA"
                    }
		            elseif($multi.formfactor -eq 22){
                    $formFactor = "FPBGA"
                    }
		            elseif($multi.formfactor -eq 23){
                    $formFactor = "LGA"
                    }
                    else{
                    $formFactor = "$formFactor [UNKNOWN]"
                    }

                $table.cell($i,1).range.text = "Form Factor"
                $table.cell($i,2).range.text = $formFactor
                $i++

                $table.cell($i,1).range.text = "Manufacturer"
                $table.cell($i,2).range.text = $multi.manufacturer
                $i++

                $table.cell($i,1).range.text = "Memory Type"

                    if($multi.memorytype -eq 0){
                    $memoryType = "Unknown"
                    }
                    elseif($multi.memorytype -eq 1){
                    $memoryType = "Other"
                    }
                    elseif($multi.memorytype -eq 2){
                    $memoryType = "DRAM"
                    }
                    elseif($multi.memorytype -eq 3){
                    $memoryType = "Synchronous DRAM"
                    }
                    elseif($multi.memorytype -eq 4){
                    $memoryType = "Cache DRAM"
                    }
                    elseif($multi.memorytype -eq 5){
                    $memoryType = "EDO"
                    }
	                elseif($multi.memorytype -eq 6){
                    $memoryType = "EDRAM"
                    }
	                elseif($multi.memorytype -eq 7){
                    $memoryType = "VRAM"
                    }
	                elseif($multi.memorytype -eq 8){
                    $memoryType = "SRAM"
                    }
	                elseif($multi.memorytype -eq 9){
                    $memoryType = "RAM"
                    }
	                elseif($multi.memorytype -eq 10){
                    $memoryType = "ROM"
                    }
	                elseif($multi.memorytype -eq 11){
                    $memoryType = "Flash"
                    }
	                elseif($multi.memorytype -eq 12){
                    $memoryType = "EEPROM"
                    }
	                elseif($multi.memorytype -eq 13){
                    $memoryType = "FEPROM"
                    }
	                elseif($multi.memorytype -eq 14){
                    $memoryType = "EPROM"
                    }
	                elseif($multi.memorytype -eq 15){
                    $memoryType = "CDRAM"
                    }
	                elseif($multi.memorytype -eq 16){
                    $memoryType = "3DRAM"
                    }
	                elseif($multi.memorytype -eq 17){
                    $memoryType = "SDRAM"
                    }
	                elseif($multi.memorytype -eq 18){
                    $memoryType = "SGRAM"
                    }
	                elseif($multi.memorytype -eq 19){
                    $memoryType = "RDRAM"
                    }
	                elseif($multi.memorytype -eq 20){
                    $memoryType = "DDR"
                    }
	                elseif($multi.memorytype -eq 21){
                    $memoryType = "DDR2"
                    }
		            elseif($multi.memorytype -eq 22){
                    $memoryType = "DDR2 FB-DIMM"
                    }
		            # No 23
		            elseif($multi.memorytype -eq 24){
                    $memoryType = "DDR3"
                    }
		            elseif($multi.memorytype -eq 25){
                    $memoryType = "FBD2"
                    }
		            else{
                    $memoryType = "$memoryType [UNKNOWN]"
                    }

                $table.cell($i,2).range.text = $memoryType
                $i++

                $table.cell($i,1).range.text = "Model"
                $table.cell($i,2).range.text = $multi.model
                $i++

                $table.cell($i,1).range.text = "Part Number"
                $table.cell($i,2).range.text = $multi.partnumber -replace " ", ""
                $i++

                $table.cell($i,1).range.text = "Serial Number"
                $table.cell($i,2).range.text = $multi.serialnumber
                $i++

                $table.cell($i,1).range.text = "Speed (Nanoseconds)"
                $table.cell($i,2).range.text = $multi.speed
                $i++

                $y++

                }

            $table.Rows.item(1).Headingformat=-1
            $table.ApplyStyleFirstColumn = $false
            $selection.InsertCaption(-2, ": [WMI] Physical Memory (Detailed)", $null, 1, $false)

            }

        }

    $selection.EndOf(15) | Out-Null
    $selection.MoveDown() | Out-Null
    $selection.InsertNewPage()
    #endregion ReportWMIPhysicalMemory

    #region ReportWMIProcessor
    # BASIC
    $selection.style = 'Heading 1'
    $selection.TypeText("[WMI] Processor")
    $selection.TypeParagraph()
    $selection.TypeParagraph()

    $selection.Style = 'Normal'
    $selection.TypeText("This data is collected from the Win32_Processor WMI class. For more information please go to https://docs.microsoft.com/en-gb/windows/desktop/CIMWin32Prov/win32-processor")
    $selection.TypeParagraph()
    $selection.TypeParagraph()

        if(($xml.info.win32_processor).count -eq 0){
        Write-Output "[WARNING] WMI Processor details not found. Moving on to next section..."
        $warnings ++
        $selection.Style = 'Normal'
        $selection.Font.Color="255"
        $selection.TypeText("[WARNING] WMI Processor data not found!")
        $selection.TypeParagraph()
        }
        elseif(($xml.info.win32_processor.errorcode) -eq 1){
        Write-Output "[WARNING] WMI Processor details not collected. Moving on to next section..."
        $warnings ++
        $selection.Style = 'Normal'
        $selection.Font.Color="255"
        $selection.TypeText("[WARNING] WMI Processor data not collected!")
        $selection.TypeParagraph()
        $selection.TypeText("Reason for error: $($xml.info.win32_processor.errortext)")
        $selection.TypeParagraph()
        }
        else{
        Write-Output "[INFO] Populating WMI Processor table."

        # Count rows in multi
        $multis = $xml.selectnodes("//info/win32_processor/win32_processor_multi")
        $count = ($multis | Measure-Object).Count

            # Calculate rows. (No of items x no of multis) + (no of multis + 1 (for header)) | or less 1 for header if single row
            if($count -eq 1){
            $rows = (25 + $count)
            }
            else{
            $rows = (25 * $count) + ($count + 1)
            }

        $table = $selection.Tables.add(
        $selection.Range,
        $rows,
        2,
        [Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior,
        [Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitWindow
        )

        $table.style = "Grid Table 4 - Accent 1"
        $table.cell(1,1).range.text = "Item"
        $table.cell(1,2).range.text = "Value"

        $i = 2 # 2 as header is row 1
        $y = 1

            foreach($multi in $multis){

                if($count -gt 1){
                $table.cell($i,1).Merge($table.cell($i, 2))
                $table.cell($i,1).range.text = "------ BLOCK $y ------"
                $i++
                }

            $table.cell($i,1).range.text = "Address Width (Bits)"
            $table.cell($i,2).range.text = $multi.addresswidth
            $i++

            $table.cell($i,1).range.text = "Caption"
            $table.cell($i,2).range.text = $multi.caption
            $i++

            $table.cell($i,1).range.text = "CPU Status"#

                if($multi.cpustatus -eq 0){
                $cPUStatus = "Unknown"
                }
                elseif($multi.cpustatus -eq 1){
                $cPUStatus = "CPU Enabled"
                }
                elseif($multi.cpustatus -eq 2){
                $cPUStatus = "CPU Disabled by User via BIOS Setup"
                }
                elseif($multi.cpustatus -eq 3){
                $cPUStatus = "CPU Disabled By BIOS (POST Error)"
                }
                elseif($multi.cpustatus -eq 4){
                $cPUStatus = "CPU is Idle"
                }
                elseif($multi.cpustatus -eq 5){
                $cPUStatus = "Reserved"
                }
	            elseif($multi.cpustatus -eq 6){
                $cPUStatus = "Reserved"
                }
	            elseif($multi.cpustatus -eq 7){
                $cPUStatus = "Other"
                }
		        else{
                $cPUStatus = "$cPUStatus [UNKNOWN]"
                }

            $table.cell($i,2).range.text = $cPUStatus
            $i++

            $table.cell($i,1).range.text = "Current Clock Speed (MHz)"
            $table.cell($i,2).range.text = $multi.currentclockspeed
            $i++

            $table.cell($i,1).range.text = "Data Width (Bits)"
            $table.cell($i,2).range.text = $multi.datawidth
            $i++

            $table.cell($i,1).range.text = "Device ID"
            $table.cell($i,2).range.text = $multi.deviceid
            $i++

            $table.cell($i,1).range.text = "Ext Clock (MHz)"
            $table.cell($i,2).range.text = $multi.extclock
            $i++

            $table.cell($i,1).range.text = "L2 Cache Size (KB)"
            $table.cell($i,2).range.text = $multi.l2cachesize
            $i++

            $table.cell($i,1).range.text = "L2 Cache Speed (MHz)"
            $table.cell($i,2).range.text = $multi.l2cachespeed
            $i++

            $table.cell($i,1).range.text = "L3 Cache Size (KB)"
            $table.cell($i,2).range.text = $multi.l3cachesize
            $i++

            $table.cell($i,1).range.text = "L3 Cache Speed (MHz)"
            $table.cell($i,2).range.text = $multi.l3cachespeed
            $i++

            $table.cell($i,1).range.text = "Load Percentage"
            $table.cell($i,2).range.text = $multi.loadpercentage
            $i++

            $table.cell($i,1).range.text = "Manufacturer"
            $table.cell($i,2).range.text = $multi.manufacturer
            $i++

            $table.cell($i,1).range.text = "Max Clock Speed (MHz)"
            $table.cell($i,2).range.text = $multi.maxclockspeed
            $i++

            $table.cell($i,1).range.text = "Name"
            $table.cell($i,2).range.text = $multi.name
            $i++

            $table.cell($i,1).range.text = "Number Of Cores"
            $table.cell($i,2).range.text = $multi.numberofcores
            $i++

            $table.cell($i,1).range.text = "Number Of Enabled Cores"
            $table.cell($i,2).range.text = $multi.numberofenabledcore
            $i++

            $table.cell($i,1).range.text = "Number Of Logical Processors"
            $table.cell($i,2).range.text = $multi.numberoflogicalprocessors
            $i++

            $table.cell($i,1).range.text = "Part Number"
            $table.cell($i,2).range.text = $multi.partnumber
            $i++

            $table.cell($i,1).range.text = "Power Management Supported"
            $table.cell($i,2).range.text = $multi.powermanagementsupported
            $i++

            $table.cell($i,1).range.text = "Processor Type"

                if($multi.processorType -eq 1){
                $processorType = "Other"
                }
                elseif($multi.processorType -eq 2){
                $processorType = "Unknown"
                }
                elseif($multi.processorType -eq 3){
                $processorType = "Central Processor"
                }
                elseif($multi.processorType -eq 4){
                $processorType = "Math Processor"
                }
                elseif($multi.processorType -eq 5){
                $processorType = "DSP Processor"
                }
	            elseif($multi.processorType -eq 6){
                $processorType = "Video Processor"
                }
		        else{
                $processorType = "$processorType [UNKNOWN]"
                }

            $table.cell($i,2).range.text = $processorType
            $i++

            $table.cell($i,1).range.text = "Serial Number"
            $table.cell($i,2).range.text = $multi.serialnumber
            $i++

            $table.cell($i,1).range.text = "Thread Count"
            $table.cell($i,2).range.text = $multi.threadcount
            $i++

            $table.cell($i,1).range.text = "VM Monitor Mode Extensions"
            $table.cell($i,2).range.text = $multi.vmmonitormodeextensions
            $i++

            $table.cell($i,1).range.text = "Virtualization Firmware Enabled"
            $table.cell($i,2).range.text = $multi.virtualizationfirmwareenabled
            $i++

            $y++

            }

        $table.Rows.item(1).Headingformat=-1
        $table.ApplyStyleFirstColumn = $false
        $selection.InsertCaption(-2, ": [WMI] Processor", $null, 1, $false)

        }

    $selection.EndOf(15) | Out-Null
    $selection.MoveDown() | Out-Null
    $selection.InsertNewPage()
    #endregion ReportWMIProcessor

    #region ReportWMILogicalDisk
    #BASIC & DETAILED
    $selection.style = 'Heading 1'
    $selection.TypeText("[WMI] Logical Disk")
    $selection.TypeParagraph()
    $selection.TypeParagraph()

    $selection.Style = 'Normal'
    $selection.TypeText("This data is collected from the Win32_LogicalDisk WMI class. For more information please go to https://docs.microsoft.com/en-gb/windows/desktop/CIMWin32Prov/win32-logicaldisk")
    $selection.TypeParagraph()
    $selection.TypeParagraph()

        if($reportType -ne 'Basic'){
        $selection.style = 'Heading 2'
        $selection.TypeText("Basic")
        $selection.TypeParagraph()
        $selection.TypeParagraph()
        }

        if(($xml.info.win32_logicaldisk).count -eq 0){
        Write-Output "[WARNING] WMI Logical Disk (Basic) details not found. Moving on to next section..."
        $warnings ++
        $selection.Style = 'Normal'
        $selection.Font.Color="255"
        $selection.TypeText("[WARNING] WMI Logical Disk (Basic) data not found!")
        $selection.TypeParagraph()
        }
        elseif(($xml.info.win32_logicaldisk.errorcode) -eq 1){
        Write-Output "[WARNING] WMI Logical Disk (Basic) details not collected. Moving on to next section..."
        $warnings ++
        $selection.Style = 'Normal'
        $selection.Font.Color="255"
        $selection.TypeText("[WARNING] WMI Logical Disk (Basic) data not collected!")
        $selection.TypeParagraph()
        $selection.TypeText("Reason for error: $($xml.info.win32_logicaldisk.errortext)")
        $selection.TypeParagraph()
        }
        else{
        Write-Output "[INFO] Populating WMI Logical Disk (Basic) table."

        # Count rows in multi
        $multis = $xml.selectnodes("//info/win32_logicaldisk/win32_logicaldisk_multi")
        $count = ($multis | Measure-Object).Count

        # Calculate rows (from query above) and add 1 for the header
        $rows = $count + 1

        $table = $selection.Tables.add(
        $selection.Range,
        $rows,
        6,
        [Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior,
        [Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitWindow
        )

        $table.style = "Grid Table 4 - Accent 1"
        $table.cell(1,1).range.text = "Device ID"
        $table.cell(1,2).range.text = "Volume Name"
        $table.cell(1,3).range.text = "Type"
        $table.cell(1,4).range.text = "Size (GB)"
        $table.cell(1,5).range.text = "Free Space (GB)"
        $table.cell(1,6).range.text = "Free (%)"

        $i = 2 # 2 as header is row 1

            foreach($multi in $multis){

            $table.cell($i,1).range.text = $multi.deviceid
            $table.cell($i,2).range.text = $multi.volumename

                if($multi.drivetype -eq 0){
                $driveType = "Unknown"
                }
                elseif($multi.drivetype -eq 1){
                $driveType = "No Root Directory"
                }
                elseif($multi.drivetype -eq 2){
                $driveType = "Removable Disk"
                }
                elseif($multi.drivetype -eq 3){
                $driveType = "Local Disk"
                }
                elseif($multi.drivetype -eq 4){
                $driveType = "Network Drive"
                }
                elseif($multi.drivetype -eq 5){
                $driveType = "Compact Disc"
                }
	            elseif($multi.drivetype -eq 6){
                $driveType = "RAM Disk"
                }
		        else{
                $driveType = "$driveType [UNKNOWN]"
                }

            $table.cell($i,3).range.text = $driveType

            # Convert bytes to GB and Round
            [string]$a1 = [system.decimal](($multi.size)/1073741824)
            $b1 = [math]::Round($a1, 2)

            $table.cell($i,4).range.text = "$b1"

            # Convert bytes to GB and Round
            [string]$a2 = [system.decimal](($multi.freespace)/1073741824)
            $b2 = [math]::Round($a2, 2)

            $table.cell($i,5).range.text = "$b2"

            # Calculate free %

                # To stop divide by zero errors, change 0 to have a very small value. It will get rounded later so the results are the same
                if($a1 -eq 0){
                $a1 = 0.000001
                }

            [string]$a3 = ($a2 / $a1) * 100
            $b3 = [math]::Round($a3, 2)

            $table.cell($i,6).range.text = "$b3"

            $i++

            }

        $table.Rows.item(1).Headingformat=-1
        $table.ApplyStyleFirstColumn = $false
        $selection.InsertCaption(-2, ": [WMI] Logical Disk (Basic)", $null, 1, $false)

          if($reportType -ne "Basic"){
          $selection.TypeParagraph()
          }

        }

    $selection.EndOf(15) | Out-Null
    $selection.MoveDown() | Out-Null

        if($reportType -eq "Detailed"){
        $selection.TypeParagraph()
        $selection.style = 'Heading 2'
        $selection.TypeText("Detailed")
        $selection.TypeParagraph()
        $selection.TypeParagraph()

            if(($xml.info.win32_logicaldisk).count -eq 0){
            Write-Output "[WARNING] WMI Logical Disk (Detailed) details not found. Moving on to next section..."
            $warnings ++
            $selection.Style = 'Normal'
            $selection.Font.Color="255"
            $selection.TypeText("[WARNING] WMI Logical Disk (Detailed) data not found!")
            $selection.TypeParagraph()
            }
            elseif(($xml.info.win32_logicaldisk.errorcode) -eq 1){
            Write-Output "[WARNING] WMI Logical Disk (Detailed) details not collected. Moving on to next section..."
            $warnings ++
            $selection.Style = 'Normal'
            $selection.Font.Color="255"
            $selection.TypeText("[WARNING] WMI Logical Disk (Detailed) data not collected!")
            $selection.TypeParagraph()
            $selection.TypeText("Reason for error: $($xml.info.win32_logicaldisk.errortext)")
            $selection.TypeParagraph()
            }
            else{
            Write-Output "[INFO] Populating WMI Logical Disk (Detailed) table."

            # Count rows in multi
            $multis = $xml.selectnodes("//info/win32_logicaldisk/win32_logicaldisk_multi")
            $count = ($multis | Measure-Object).Count

                # Calculate rows. (No of items x no of multis) + (no of multis + 1 (for header)) | or less 1 for header if single row
                if($count -eq 1){
                $rows = (9 + $count)
                }
                else{
                $rows = (9 * $count) + ($count + 1)
                }

            $table = $selection.Tables.add(
            $selection.Range,
            $rows,
            2,
            [Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior,
            [Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitWindow
            )

            $table.style = "Grid Table 4 - Accent 1"
            $table.cell(1,1).range.text = "Item"
            $table.cell(1,2).range.text = "Value"

            $i = 2 # 2 as header is row 1
            $y = 1

                foreach($multi in $multis){

                    if($count -gt 1){
                    $table.cell($i,1).Merge($table.cell($i, 2))
                    $table.cell($i,1).range.text = "------ BLOCK $y ------"
                    $i++
                    }

                $table.cell($i,1).range.text = "Device ID"
                $table.cell($i,2).range.text = $multi.deviceid
                $i++

                $table.cell($i,1).range.text = "Description"
                $table.cell($i,2).range.text = $multi.description
                $i++

                $table.cell($i,1).range.text = "Drive Type"

        	        if($multi.drivetype -eq 0){
                    $driveType = "Unknown"
                    }
                    elseif($multi.drivetype -eq 1){
                    $driveType = "No Root Directory"
                    }
                    elseif($multi.drivetype -eq 2){
                    $driveType = "Removable Disk"
                    }
                    elseif($multi.drivetype -eq 3){
                    $driveType = "Local Disk"
                    }
                    elseif($multi.drivetype -eq 4){
                    $driveType = "Network Drive"
                    }
                    elseif($multi.drivetype -eq 5){
                    $driveType = "Compact Disc"
                    }
	                elseif($multi.drivetype -eq 6){
                    $driveType = "RAM Disk"
                    }
		            else{
                    $driveType = "$driveType [UNKNOWN]"
                    }

                $table.cell($i,2).range.text = $driveType
                $i++

                $table.cell($i,1).range.text = "File System"
                $table.cell($i,2).range.text = $multi.filesystem
                $i++

                $table.cell($i,1).range.text = "Free Space (GB)"

                # Convert bytes to GB and Round
                [string]$a = [system.decimal](($multi.freespace)/1073741824)
                $b = [math]::Round($a, 2)

                $table.cell($i,2).range.text = "$b"
                $i++

                $table.cell($i,1).range.text = "Media Type"

        	        if($multi.mediatype -eq 0){
                    $mediaType = "Format is unknown"
                    }
			        elseif(($multi.mediatype -ge 1) -and ($multi.mediatype -le 10)){
                    $mediaType = "Floppy Disk"
                    }
                    elseif($multi.mediatype -eq 11){
                    $mediaType = "Removable media other than floppy"
                    }
                    elseif($multi.mediatype -eq 12){
                    $mediaType = "Fixed hard disk media"
                    }
                    elseif(($multi.mediatype -ge 13) -and ($multi.mediatype -le 22)){
                    $mediaType = "Floppy Disk"
                    }
                    else{
                    $mediaType = "$mediaType [UNKNOWN]"
                    }

                $table.cell($i,2).range.text = $mediaType
                $i++

                $table.cell($i,1).range.text = "Size (GB)"

                # Convert bytes to GB and Round
                [string]$a = [system.decimal](($multi.size)/1073741824)
                $b = [math]::Round($a, 2)

                $table.cell($i,2).range.text = "$b"
                $i++

                $table.cell($i,1).range.text = "Volume Name"
                $table.cell($i,2).range.text = $multi.volumename
                $i++

                $table.cell($i,1).range.text = "Volume Serial Number"
                $table.cell($i,2).range.text = $multi.volumeserialnumber
                $i++

                $y++

                }

            $table.Rows.item(1).Headingformat=-1
            $table.ApplyStyleFirstColumn = $false
            $selection.InsertCaption(-2, ": [WMI] Logical Disk (Detailed)", $null, 1, $false)

            }

        }

    $selection.EndOf(15) | Out-Null
    $selection.MoveDown() | Out-Null
    $selection.InsertNewPage()
    #endregion ReportWMILogicalDisk

    #region ReportWMIDiskPartition
    # BASIC & DETAILED
    $selection.style = 'Heading 1'
    $selection.TypeText("[WMI] Disk Partition")
    $selection.TypeParagraph()
    $selection.TypeParagraph()

    $selection.Style = 'Normal'
    $selection.TypeText("This data is collected from the Win32_DiskPartition WMI class. For more information please go to https://docs.microsoft.com/en-gb/windows/desktop/CIMWin32Prov/win32-diskpartition")
    $selection.TypeParagraph()
    $selection.TypeParagraph()

        if($reportType -ne 'Basic'){
        $selection.style = 'Heading 2'
        $selection.TypeText("Basic")
        $selection.TypeParagraph()
        $selection.TypeParagraph()
        }

        if(($xml.info.win32_diskpartition).count -eq 0){
        Write-Output "[WARNING] WMI - Disk Partition (Basic) details not found. Moving on to next section..."
        $warnings ++
        $selection.Style = 'Normal'
        $selection.Font.Color="255"
        $selection.TypeText("[WARNING] WMI Disk Partition (Basic) data not found!")
        $selection.TypeParagraph()
        }
        elseif(($xml.info.win32_diskpartition.errorcode) -eq 1){
        Write-Output "[WARNING] WMI Disk Partition (Basic) details not collected. Moving on to next section..."
        $warnings ++
        $selection.Style = 'Normal'
        $selection.Font.Color="255"
        $selection.TypeText("[WARNING] WMI Disk Partition (Basic) data not collected!")
        $selection.TypeParagraph()
        $selection.TypeText("Reason for error: $($xml.info.win32_diskpartition.errortext)")
        $selection.TypeParagraph()
        }
        else{
        Write-Output "[INFO] Populating WMI Disk Partition (Basic) table."

        # Count rows in multi
        $multis = $xml.selectnodes("//info/win32_diskpartition/win32_diskpartition_multi")
        $count = ($multis | Measure-Object).Count

        # Calculate rows (from query above) and add 1 for the header
        $rows = $count + 1

        $table = $selection.Tables.add(
        $selection.Range,
        $rows,
        6,
        [Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior,
        [Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitWindow
        )

        $table.style = "Grid Table 4 - Accent 1"
        $table.cell(1,1).range.text = "Device ID"
        $table.cell(1,2).range.text = "Bootable"
        $table.cell(1,3).range.text = "Block Size (Bytes)"
        $table.cell(1,4).range.text = "Disk Index"
        $table.cell(1,5).range.text = "Primary Partition"
        $table.cell(1,6).range.text = "Size (GB)"

        $i = 2 # 2 as header is row 1

            foreach($multi in $multis){

            $table.cell($i,1).range.text = $multi.deviceid
            $table.cell($i,2).range.text = $multi.bootable
            $table.cell($i,3).range.text = $multi.blocksize
            $table.cell($i,4).range.text = $multi.diskindex
            $table.cell($i,5).range.text = $multi.primarypartition

            # Convert bytes to GB and Round
            [string]$a = [system.decimal](($multi.size)/1073741824)
            $b = [math]::Round($a,2)

            $table.cell($i,6).range.text = "$b"
            $i++

            }

        $table.Rows.item(1).Headingformat=-1
        $table.ApplyStyleFirstColumn = $false
        $selection.InsertCaption(-2, ": [WMI] Disk Partition (Basic)", $null, 1, $false)

          if($reportType -ne "Basic"){
          $selection.TypeParagraph()
          }

        }

    $selection.EndOf(15) | Out-Null
    $selection.MoveDown() | Out-Null

        if($reportType -eq "Detailed"){
        $selection.TypeParagraph()
        $selection.style = 'Heading 2'
        $selection.TypeText("Detailed")
        $selection.TypeParagraph()
        $selection.TypeParagraph()

            if(($xml.info.win32_diskpartition).count -eq 0){
            Write-Output "[WARNING] WMI Disk Partition (Detailed) details not found. Moving on to next section..."
            $warnings ++
            $selection.Style = 'Normal'
            $selection.Font.Color="255"
            $selection.TypeText("[WARNING] WMI Disk Partition (Detailed) data not found!")
            $selection.TypeParagraph()
            }
            elseif(($xml.info.win32_diskpartition.errorcode) -eq 1){
            Write-Output "[WARNING] WMI Disk Partition (Detailed) details not collected. Moving on to next section..."
            $warnings ++
            $selection.Style = 'Normal'
            $selection.Font.Color="255"
            $selection.TypeText("[WARNING] WMI Disk Partition (Detailed) data not collected!")
            $selection.TypeParagraph()
            $selection.TypeText("Reason for error: $($xml.info.win32_diskpartition.errortext)")
            $selection.TypeParagraph()
            }
            else{
            Write-Output "[INFO] Populating WMI Disk Partition (Detailed) table."

            # Count rows in multi
            $multis = $xml.selectnodes("//info/win32_diskpartition/win32_diskpartition_multi")
            $count = ($multis | Measure-Object).Count

                # Calculate rows. (No of items x no of multis) + (no of multis + 1 (for header)) | or less 1 for header if single row
                if($count -eq 1){
                $rows = (9 + $count)
                }
                else{
                $rows = (9 * $count) + ($count + 1)
                }

            $table = $selection.Tables.add(
            $selection.Range,
            $rows,
            2,
            [Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior,
            [Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitWindow
            )

            $table.style = "Grid Table 4 - Accent 1"
            $table.cell(1,1).range.text = "Item"
            $table.cell(1,2).range.text = "Value"

            $i = 2 # 2 as header is row 1
            $y = 1

                foreach($multi in $multis){

                    if($count -gt 1){
                    $table.cell($i,1).Merge($table.cell($i, 2))
                    $table.cell($i,1).range.text = "------ BLOCK $y ------"
                    $i++
                    }

                $table.cell($i,1).range.text = "Block Size (Bytes)"
                $table.cell($i,2).range.text = $multi.blocksize
                $i++

                $table.cell($i,1).range.text = "Bootable"
                $table.cell($i,2).range.text = $multi.bootable
                $i++

                $table.cell($i,1).range.text = "Boot Partition"
                $table.cell($i,2).range.text = $multi.bootpartition
                $i++

                $table.cell($i,1).range.text = "Description"
                $table.cell($i,2).range.text = $multi.description
                $i++

                $table.cell($i,1).range.text = "Device ID"
                $table.cell($i,2).range.text = $multi.deviceid
                $i++

                $table.cell($i,1).range.text = "Disk Index"
                $table.cell($i,2).range.text = $multi.diskindex
                $i++

                $table.cell($i,1).range.text = "Number Of Blocks"
                $table.cell($i,2).range.text = $multi.numberofblocks
                $i++

                $table.cell($i,1).range.text = "Primary Partition"
                $table.cell($i,2).range.text = $multi.primarypartition
                $i++

                # Convert bytes to GB and Round
                [string]$a = [system.decimal](($multi.size)/1073741824)
                $b = [math]::Round($a,2)

                $table.cell($i,1).range.text = "Size (GB)"
                $table.cell($i,2).range.text = "$b"
                $i++

                $y++

                }

            $table.Rows.item(1).Headingformat=-1
            $table.ApplyStyleFirstColumn = $false
            $selection.InsertCaption(-2, ": [WMI] Disk Partition (Detailed)", $null, 1, $false)

            }

        }

    $selection.EndOf(15) | Out-Null
    $selection.MoveDown() | Out-Null
    $selection.InsertNewPage()
    #endregion ReportWMIDiskPartition

    #region ReportWMIShare
    # BASIC
    $selection.style = 'Heading 1'
    $selection.TypeText("[WMI] Share")
    $selection.TypeParagraph()
    $selection.TypeParagraph()

    $selection.Style = 'Normal'
    $selection.TypeText("This data is collected from the Win32_Share WMI class. For more information please go to https://docs.microsoft.com/en-gb/windows/desktop/CIMWin32Prov/win32-share")
    $selection.TypeParagraph()
    $selection.TypeParagraph()

        if(($xml.info.win32_share).count -eq 0){
        Write-Output "[WARNING] WMI Share data not found. Moving on to next section..."
        $warnings ++
        $selection.Style = 'Normal'
        $selection.Font.Color="255"
        $selection.TypeText("[WARNING] WMI Share data not found!")
        $selection.TypeParagraph()
        }
        elseif(($xml.info.win32_share.errorcode) -eq 1){
        Write-Output "[WARNING] WMI Share details not collected. Moving on to next section..."
        $warnings ++
        $selection.Style = 'Normal'
        $selection.Font.Color="255"
        $selection.TypeText("[WARNING] WMI Share data not collected!")
        $selection.TypeParagraph()
        $selection.TypeText("Reason for error: $($xml.info.win32_share.errortext)")
        $selection.TypeParagraph()
        }
        else{
        Write-Output "[INFO] Populating WMI Share table."

        # Count rows in multi
        $multis = $xml.selectnodes("//info/win32_share/win32_share_multi")
        $count = ($multis | Measure-Object).Count

            # Calculate rows (from query above) and add 1 for the header
            if($count -eq 0){
            $rows = $count + 2
            }
            else{
            $rows = $count + 1
            }

        $table = $selection.Tables.add(
        $selection.Range,
        $rows,
        5,
        [Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior,
        [Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitWindow
        )

        $table.style = "Grid Table 4 - Accent 1"
        $table.cell(1,1).range.text = "Name"
        $table.cell(1,2).range.text = "Caption"
        $table.cell(1,3).range.text = "Description"
        $table.cell(1,4).range.text = "Path"
        $table.cell(1,5).range.text = "Type"

        $i = 2 # 2 as header is row 1

            if($count -eq 0){
            $table.cell($i,1).range.text = "No shares found"
            $table.Cell($i,2).Merge($table.Cell($i,1))
            $table.Cell($i,2).Merge($table.Cell($i,1))
            $table.Cell($i,2).Merge($table.Cell($i,1))
            $table.Cell($i,2).Merge($table.Cell($i,1))
            }

            foreach($multi in $multis){

            $table.cell($i,1).range.text = $multi.name
            $table.cell($i,2).range.text = $multi.caption
            $table.cell($i,3).range.text = $multi.description
            $table.cell($i,4).range.text = $multi.path

                if($multi.type -eq 0){
                $type = "Disk Drive"
                }
                elseif($multi.type -eq 1){
                $type = "Print Queue"
                }
                elseif($multi.type -eq 2){
                $type = "Device"
                }
                elseif($multi.type -eq 3){
                $type = "IPC"
                }
                elseif($multi.type -eq 2147483648){
                $type = "Disk Drive Admin"
                }
                elseif($multi.type -eq 2147483649){
                $type = "Print Queue Admin"
                }
	            elseif($multi.type -eq 2147483650){
                $type = "Device Admin"
                }
	            elseif($multi.type -eq 2147483651){
                $type = "IPC Admin"
                }
		        else{
                $type = "$type [UNKNOWN]"
                }

            $table.cell($i,5).range.text = $type
            $i++

            }

        $table.Rows.item(1).Headingformat=-1
        $table.ApplyStyleFirstColumn = $false
        $selection.InsertCaption(-2, ": [WMI] Share", $null, 1, $false)

        }

    $selection.EndOf(15) | Out-Null
    $selection.MoveDown() | Out-Null
    $selection.InsertNewPage()
    #endregion ReportWMIShare

    #region ReportWMIStartUpCommand
    # BASIC & DETAILED
    $selection.style = 'Heading 1'
    $selection.TypeText("[WMI] Start Up Command")
    $selection.TypeParagraph()
    $selection.TypeParagraph()

    $selection.Style = 'Normal'
    $selection.TypeText("This data is collected from the Win32_StartUpCommand WMI class. For more information please go to https://docs.microsoft.com/en-gb/windows/desktop/CIMWin32Prov/win32-startupcommand")
    $selection.TypeParagraph()
    $selection.TypeParagraph()

        if($reportType -ne 'Basic'){
        $selection.style = 'Heading 2'
        $selection.TypeText("Basic")
        $selection.TypeParagraph()
        $selection.TypeParagraph()
        }

        if(($xml.info.win32_startupcommand).count -eq 0){
        Write-Output "[WARNING] WMI Start Up Command (Basic) details not found. Moving on to next section..."
        $warnings ++
        $selection.Style = 'Normal'
        $selection.Font.Color="255"
        $selection.TypeText("[WARNING] WMI Start Up Command (Basic) data not found!")
        $selection.TypeParagraph()
        }
        elseif(($xml.info.win32_startupcommand.errorcode) -eq 1){
        Write-Output "[WARNING] WMI Start Up Command (Basic) details not collected. Moving on to next section..."
        $warnings ++
        $selection.Style = 'Normal'
        $selection.Font.Color="255"
        $selection.TypeText("[WARNING] WMI Start Up Command (Basic) data not collected!")
        $selection.TypeParagraph()
        $selection.TypeText("Reason for error: $($xml.info.win32_startupcommand.errortext)")
        $selection.TypeParagraph()
        }
        else{
        Write-Output "[INFO] Populating WMI Start Up Command (Basic) table."

        # Count rows in multi
        $multis = $xml.selectnodes("//info/win32_startupcommand/win32_startupcommand_multi")
        $count = ($multis | Measure-Object).Count

            # Calculate rows (from query above) and add 1 for the header
            if($count -eq 0){
            $rows = $count + 2
            }
            else{
            $rows = $count + 1
            }

        $table = $selection.Tables.add(
        $selection.Range,
        $rows,
        3,
        [Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior,
        [Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitWindow
        )

        $table.style = "Grid Table 4 - Accent 1"
        $table.cell(1,1).range.text = "Name"
        $table.cell(1,2).range.text = "Command"
        $table.cell(1,3).range.text = "User"

        $i = 2 # 2 as header is row 1

            if($count -eq 0){
            $table.cell($i,1).range.text = "No start up commands configured"
            $table.Cell($i,2).Merge($table.Cell($i,1))
            $table.Cell($i,2).Merge($table.Cell($i,1))
            }

            foreach($multi in $multis){

            $table.cell($i,1).range.text = $multi.name
            $table.cell($i,2).range.text = $multi.command
            $table.cell($i,3).range.text = $multi.user
            $i++

            }

        $table.Rows.item(1).Headingformat=-1
        $table.ApplyStyleFirstColumn = $false
        $selection.InsertCaption(-2, ": [WMI] Start Up Command (Basic)", $null, 1, $false)

          if($reportType -ne "Basic"){
          $selection.TypeParagraph()
          }

        }

    $selection.EndOf(15) | Out-Null
    $selection.MoveDown() | Out-Null

        if($reportType -eq "Detailed"){
        $selection.TypeParagraph()
        $selection.style = 'Heading 2'
        $selection.TypeText("Detailed")
        $selection.TypeParagraph()
        $selection.TypeParagraph()

            if(($xml.info.win32_startupcommand).count -eq 0){
            Write-Output "[WARNING] WMI Start Up Command (Detailed) details not found. Moving on to next section..."
            $warnings ++
            $selection.Style = 'Normal'
            $selection.Font.Color="255"
            $selection.TypeText("[WARNING] WMI Start Up Command (Detailed) data not found!")
            $selection.TypeParagraph()
            }
            elseif(($xml.info.win32_startupcommand.errorcode) -eq 1){
            Write-Output "[WARNING] WMI Start Up Command (Detailed) details not collected. Moving on to next section..."
            $warnings ++
            $selection.Style = 'Normal'
            $selection.Font.Color="255"
            $selection.TypeText("[WARNING] WMI Start Up Command (Detailed) data not collected!")
            $selection.TypeParagraph()
            $selection.TypeText("Reason for error: $($xml.info.win32_startupcommand.errortext)")
            $selection.TypeParagraph()
            }
            else{
            Write-Output "[INFO] Populating WMI Start Up Command (Detailed) table."

            # Count rows in multi
            $multis = $xml.selectnodes("//info/win32_startupcommand/win32_startupcommand_multi")
            $count = ($multis | Measure-Object).Count

                # Calculate rows. (No of items x no of multis) + (no of multis + 1 (for header)) | or less 1 for header if single row
                if($count -eq 0){
                $rows = $count + 2
                }
                elseif($count -eq 1){
                $rows = (5 + $count)
                }
                else{
                $rows = (5 * $count) + ($count + 1)
                }

            $table = $selection.Tables.add(
            $selection.Range,
            $rows,
            2,
            [Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior,
            [Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitWindow
            )

            $table.style = "Grid Table 4 - Accent 1"
            $table.cell(1,1).range.text = "Item"
            $table.cell(1,2).range.text = "Value"

            $i = 2 # 2 as header is row 1
            $y = 1

                if($count -eq 0){
                $table.cell($i,1).range.text = "No start up commands configured"
                $table.Cell($i,2).Merge($table.Cell($i,1))
                }

                foreach($multi in $multis){

                    if($count -gt 1){
                    $table.cell($i,1).Merge($table.cell($i, 2))
                    $table.cell($i,1).range.text = "------ BLOCK $y ------"
                    $i++
                    }

                $table.cell($i,1).range.text = "Caption"
                $table.cell($i,2).range.text = $multi.caption
                $i++

                $table.cell($i,1).range.text = "Command"
                $table.cell($i,2).range.text = $multi.command
                $i++

                $table.cell($i,1).range.text = "Description"
                $table.cell($i,2).range.text = $multi.description
                $i++

                $table.cell($i,1).range.text = "Name"
                $table.cell($i,2).range.text = $multi.name
                $i++

                $table.cell($i,1).range.text = "User"
                $table.cell($i,2).range.text = $multi.user
                $i++

                $y++

                }

            $table.Rows.item(1).Headingformat=-1
            $table.ApplyStyleFirstColumn = $false
            $selection.InsertCaption(-2, ": [WMI] Start Up Command (Detailed)", $null, 1, $false)

            }

        }

    $selection.EndOf(15) | Out-Null
    $selection.MoveDown() | Out-Null
    $selection.InsertNewPage()
    #endregion ReportWMIStartUpCommand

    #region ReportWMIPageFileUsage
    # BASIC
    $selection.style = 'Heading 1'
    $selection.TypeText("[WMI] Page File Usage")
    $selection.TypeParagraph()
    $selection.TypeParagraph()

    $selection.Style = 'Normal'
    $selection.TypeText("This data is collected from the Win32_PageFileUsage WMI class. For more information please go to https://docs.microsoft.com/en-gb/windows/desktop/CIMWin32Prov/win32-pagefileusage")
    $selection.TypeParagraph()
    $selection.TypeParagraph()

        if(($xml.info.win32_pagefileusage).count -eq 0){
        Write-Output "[WARNING] WMI Page File Usage details not found. Moving on to next section..."
        $warnings ++
        $selection.Style = 'Normal'
        $selection.Font.Color="255"
        $selection.TypeText("[WARNING] WMI Page File Usage data not found!")
        $selection.TypeParagraph()
        }
        elseif(($xml.info.win32_pagefileusage.errorcode) -eq 1){
        Write-Output "[WARNING] WMI Page File Usage details not collected. Moving on to next section..."
        $warnings ++
        $selection.Style = 'Normal'
        $selection.Font.Color="255"
        $selection.TypeText("[WARNING] WMI Page File Usage data not collected!")
        $selection.TypeParagraph()
        $selection.TypeText("Reason for error: $($xml.info.win32_pagefileusage.errortext)")
        $selection.TypeParagraph()
        }
        else{
        Write-Output "[INFO] Populating WMI Page File Usage table."

        # Count rows in multi
        $multis = $xml.selectnodes("//info/win32_pagefileusage/win32_pagefileusage_multi")
        $count = ($multis | Measure-Object).Count

            # Calculate rows. (No of items x no of multis) + (no of multis + 1 (for header)) | or less 1 for header if single row
            if($count -eq 0){
            $rows = $count + 2
            }
            elseif($count -eq 1){
            $rows = (5 + $count)
            }
            else{
            $rows = (5 * $count) + ($count + 1)
            }

        $table = $selection.Tables.add(
        $selection.Range,
        $rows,
        2,
        [Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior,
        [Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitWindow
        )

        $table.style = "Grid Table 4 - Accent 1"
        $table.cell(1,1).range.text = "Item"
        $table.cell(1,2).range.text = "Value"

        $i = 2 # 2 as header is row 1
        $y = 1

            if($count -eq 0){
            $table.cell($i,1).range.text = "No page file configured"
            $table.Cell($i,2).Merge($table.Cell($i,1))
            }

            foreach($multi in $multis){

                if($count -gt 1){
                $table.cell($i,1).Merge($table.cell($i, 2))
                $table.cell($i,1).range.text = "------ BLOCK $y ------"
                $i++
                }

            $table.cell($i,1).range.text = "Current Usage (GB)"

            # Megabtyes to Gigabytes
            $a = [system.decimal](($multi.currentusage)/1000)
            $b = [math]::Round($a, 2)

            $table.cell($i,2).range.text = "$b"
            $i++

            $table.cell($i,1).range.text = "Allocated Base Size (GB)"

            # Megabtyes to Gigabytes
            $a = [system.decimal](($multi.allocatedbasesize)/1000)
            $b = [math]::Round($a, 2)

            $table.cell($i,2).range.text = "$b"
            $i++

            $table.cell($i,1).range.text = "Caption"
            $table.cell($i,2).range.text = $multi.caption
            $i++

            $table.cell($i,1).range.text = "Description"
            $table.cell($i,2).range.text = $multi.description
            $i++

            $table.cell($i,1).range.text = "Name"
            $table.cell($i,2).range.text = $multi.name
            $i++

            $y++

            }

        $table.Rows.item(1).Headingformat=-1
        $table.ApplyStyleFirstColumn = $false
        $selection.InsertCaption(-2, ": [WMI] Page File Usage", $null, 1, $false)

        }

    $selection.EndOf(15) | Out-Null
    $selection.MoveDown() | Out-Null
    $selection.InsertNewPage()
    #endregion ReportWMIPageFileUsage

    #region ReportWMIQuickFixEngineering
    # BAISC & DETAILED
    $selection.style = 'Heading 1'
    $selection.TypeText("[WMI] Quick Fix Engineering (Updates)")
    $selection.TypeParagraph()
    $selection.TypeParagraph()

    $selection.Style = 'Normal'
    $selection.TypeText("This data is collected from the Win32_QuickFixEngineering WMI class. For more information please go to https://docs.microsoft.com/en-gb/windows/desktop/CIMWin32Prov/win32-quickfixengineering")
    $selection.TypeParagraph()
    $selection.TypeParagraph()

        if($reportType -ne 'Basic'){
        $selection.style = 'Heading 2'
        $selection.TypeText("Basic")
        $selection.TypeParagraph()
        $selection.TypeParagraph()
        }

        if(($xml.info.win32_quickfixengineering).count -eq 0){
        Write-Output "[WARNING] WMI Quick Fix Engineering (Basic) details not found. Moving on to next section..."
        $warnings ++
        $selection.Style = 'Normal'
        $selection.Font.Color="255"
        $selection.TypeText("[WARNING] WMI Quick Fix Engineering (Basic) data not found!")
        $selection.TypeParagraph()
        }
        elseif(($xml.info.win32_quickfixengineering.errorcode) -eq 1){
        Write-Output "[WARNING] WMI Quick Fix Engineering (Basic) details not collected. Moving on to next section..."
        $warnings ++
        $selection.Style = 'Normal'
        $selection.Font.Color="255"
        $selection.TypeText("[WARNING] WMI Quick Fix Engineering (Basic) data not collected!")
        $selection.TypeParagraph()
        $selection.TypeText("Reason for error: $($xml.info.win32_quickfixengineering.errortext)")
        $selection.TypeParagraph()
        }
        else{
        Write-Output "[INFO] Populating WMI Quick Fix Engineering (Basic) table."

        # Count rows in multi
        $multis = $xml.selectnodes("//info/win32_quickfixengineering/win32_quickfixengineering_multi")
        $count = ($multis | Measure-Object).Count

            # Calculate rows (from query above) and add 1 for the header
            if($count -eq 0){
            $rows = $count + 2
            }
            else{
            $rows = $count + 1
            }

        $table = $selection.Tables.add(
        $selection.Range,
        $rows,
        3,
        [Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior,
        [Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitWindow
        )

        $table.style = "Grid Table 4 - Accent 1"
        $table.cell(1,1).range.text = "Hot Fix ID"
        $table.cell(1,2).range.text = "Description"
        $table.cell(1,3).range.text = "Installed On"

        $i = 2 # 2 as header is row 1

            if($count -eq 0){
            $table.cell($i,1).range.text = "No updates installed"
            $table.Cell($i,2).Merge($table.Cell($i,1))
            $table.Cell($i,2).Merge($table.Cell($i,1))
            }

            foreach($multi in $multis){

            $table.cell($i,1).range.text = $multi.hotfixid
            $table.cell($i,2).range.text = $multi.description

                if($multi.installedon -eq ''){
                $instOn = ""
                }
                else{
                $instOn =
                $multi.installedon.Substring(3,2) + "/" +
                $multi.installedon.Substring(0,2) + "/" +
                $multi.installedon.Substring(6,4)
                }

            $table.cell($i,3).range.text = $instOn
            $i++

            }

        $table.Rows.item(1).Headingformat=-1
        $table.ApplyStyleFirstColumn = $false
        $selection.InsertCaption(-2, ": [WMI] Quick Fix Engineering (Basic)", $null, 1, $false)

          if($reportType -ne "Basic"){
          $selection.TypeParagraph()
          }

        }

    $selection.EndOf(15) | Out-Null
    $selection.MoveDown() | Out-Null

        if($reportType -eq "Detailed"){
        $selection.TypeParagraph()
        $selection.style = 'Heading 2'
        $selection.TypeText("Detailed")
        $selection.TypeParagraph()
        $selection.TypeParagraph()

            if(($xml.info.win32_quickfixengineering).count -eq 0){
            Write-Output "[WARNING] WMI Quick Fix Engineering (Detailed) details not found. Moving on to next section..."
            $warnings ++
            $selection.Style = 'Normal'
            $selection.Font.Color="255"
            $selection.TypeText("[WARNING] WMI Quick Fix Engineering (Detailed) data not found!")
            $selection.TypeParagraph()
            }
            elseif(($xml.info.win32_quickfixengineering.errorcode) -eq 1){
            Write-Output "[WARNING] WMI Quick Fix Engineering (Detailed) details not collected. Moving on to next section..."
            $warnings ++
            $selection.Style = 'Normal'
            $selection.Font.Color="255"
            $selection.TypeText("[WARNING] WMI Quick Fix Engineering (Detailed) data not collected!")
            $selection.TypeParagraph()
            $selection.TypeText("Reason for error: $($xml.info.win32_quickfixengineering.errortext)")
            $selection.TypeParagraph()
            }
            else{
            Write-Output "[INFO] Populating WMI Quick Fix Engineering (Detailed) table."

            # Count rows in multi
            $multis = $xml.selectnodes("//info/win32_quickfixengineering/win32_quickfixengineering_multi")
            $count = ($multis | Measure-Object).Count

                # Calculate rows. (No of items x no of multis) + (no of multis + 1 (for header)) | or less 1 for header if single row
                if($count -eq 0){
                $rows = $count + 2
                }
                elseif($count -eq 1){
                $rows = (5 + $count)
                }
                else{
                $rows = (5 * $count) + ($count + 1)
                }

            $table = $selection.Tables.add(
            $selection.Range,
            $rows,
            2,
            [Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior,
            [Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitWindow
            )

            $table.style = "Grid Table 4 - Accent 1"
            $table.cell(1,1).range.text = "Item"
            $table.cell(1,2).range.text = "Value"

            $i = 2 # 2 as header is row 1
            $y = 1

                if($count -eq 0){
                $table.cell($i,1).range.text = "No programs installed"
                $table.Cell($i,2).Merge($table.Cell($i,1))
                }

                foreach($multi in $multis){

                    if($count -gt 1){
                    $table.cell($i,1).Merge($table.cell($i, 2))
                    $table.cell($i,1).range.text = "------ BLOCK $y ------"
                    $i++
                    }

                $table.cell($i,1).range.text = "Caption"
                $table.cell($i,2).range.text = $multi.caption
                $i++

                $table.cell($i,1).range.text = "Description"
                $table.cell($i,2).range.text = $multi.description
                $i++

                $table.cell($i,1).range.text = "Hot Fix ID"
                $table.cell($i,2).range.text = $multi.hotfixid
                $i++

                $table.cell($i,1).range.text = "Installed By"
                $table.cell($i,2).range.text = $multi.installedby
                $i++

                $table.cell($i,1).range.text = "Installed On"

                    if($multi.installedon -eq ''){
                    $instOn = ""
                    }
                    else{

                    $instOn =
                    $multi.installedon.Substring(3,2) + "/" +
                    $multi.installedon.Substring(0,2) + "/" +
                    $multi.installedon.Substring(6,4)
                    }

                $table.cell($i,2).range.text = $instOn
                $i++

                $y++

                }

            $table.Rows.item(1).Headingformat=-1
            $table.ApplyStyleFirstColumn = $false
            $selection.InsertCaption(-2, ": [WMI] Quick Fix Engineering (Detailed)", $null, 1, $false)

            }

        }

    $selection.EndOf(15) | Out-Null
    $selection.MoveDown() | Out-Null
    $selection.InsertNewPage()
    #endregion ReportWMIQuickFixEngineering

    #region ReportWMISystemEnclosure
    # BASIC
    $selection.style = 'Heading 1'
    $selection.TypeText("[WMI] System Enclosure")
    $selection.TypeParagraph()
    $selection.TypeParagraph()

    $selection.Style = 'Normal'
    $selection.TypeText("This data is collected from the Win32_SystemEnclosure WMI class. For more information please go to https://docs.microsoft.com/en-gb/windows/desktop/CIMWin32Prov/win32-systemenclosure")
    $selection.TypeParagraph()
    $selection.TypeParagraph()

        if(($xml.info.win32_systemenclosure).count -eq 0){
        Write-Output "[WARNING] WMI System Enclosure details not found. Moving on to next section..."
        $warnings ++
        $selection.Style = 'Normal'
        $selection.Font.Color="255"
        $selection.TypeText("[WARNING] WMI System Enclosure data not found!")
        $selection.TypeParagraph()
        }
        elseif(($xml.info.win32_systemenclosure.errorcode) -eq 1){
        Write-Output "[WARNING] WMI System Enclosure details not collected. Moving on to next section..."
        $warnings ++
        $selection.Style = 'Normal'
        $selection.Font.Color="255"
        $selection.TypeText("[WARNING] WMI System Enclosure data not collected!")
        $selection.TypeParagraph()
        $selection.TypeText("Reason for error: $($xml.info.win32_systemenclosure.errortext)")
        $selection.TypeParagraph()
        }
        else{
        Write-Output "[INFO] Populating WMI System Enclosure table."

        $table = $selection.Tables.add(
        $selection.Range,
        7,
        2,
        [Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior,
        [Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitWindow
        )

        $table.style = "Grid Table 4 - Accent 1"
        $table.cell(1,1).range.text = "Item"
        $table.cell(1,2).range.text = "Value"

        $table.cell(2,1).range.text = "Chassis Types"

            if($xml.info.win32_systemenclosure.chassistypes -eq 1){
            $chassisTypes = "Other"
            }
            elseif($xml.info.win32_systemenclosure.chassistypes -eq 2){
            $chassisTypes = "Unknown"
            }
            elseif($xml.info.win32_systemenclosure.chassistypes -eq 3){
            $chassisTypes = "Desktop"
            }
            elseif($xml.info.win32_systemenclosure.chassistypes -eq 4){
            $chassisTypes = "Low Profile Desktop"
            }
            elseif($xml.info.win32_systemenclosure.chassistypes -eq 5){
            $chassisTypes = "Pizza Box"
            }
            elseif($xml.info.win32_systemenclosure.chassistypes -eq 6){
            $chassisTypes = "Mini Tower"
            }
            elseif($xml.info.win32_systemenclosure.chassistypes -eq 7){
            $chassisTypes = "Tower"
            }
            elseif($xml.info.win32_systemenclosure.chassistypes -eq 8){
            $chassisTypes = "Portable"
            }
            elseif($xml.info.win32_systemenclosure.chassistypes -eq 9){
            $chassisTypes = "Laptop"
            }
            elseif($xml.info.win32_systemenclosure.chassistypes -eq 10){
            $chassisTypes = "Notebook"
            }
            elseif($xml.info.win32_systemenclosure.chassistypes -eq 11){
            $chassisTypes = "Hand Held"
            }
            elseif($xml.info.win32_systemenclosure.chassistypes -eq 12){
            $chassisTypes = "Docking Station"
            }
            elseif($xml.info.win32_systemenclosure.chassistypes -eq 13){
            $chassisTypes = "All In One"
            }
            elseif($xml.info.win32_systemenclosure.chassistypes -eq 14){
            $chassisTypes = "Sub Notebook"
            }
            elseif($xml.info.win32_systemenclosure.chassistypes -eq 15){
            $chassisTypes = "Space-Saving"
            }
            elseif($xml.info.win32_systemenclosure.chassistypes -eq 16){
            $chassisTypes = "Lunch Box"
            }
            elseif($xml.info.win32_systemenclosure.chassistypes -eq 17){
            $chassisTypes = "Main System Chassis"
            }
            elseif($xml.info.win32_systemenclosure.chassistypes -eq 18){
            $chassisTypes = "Expansion Chassis"
            }
            elseif($xml.info.win32_systemenclosure.chassistypes -eq 19){
            $chassisTypes = "Sub Chassis"
            }
		    elseif($xml.info.win32_systemenclosure.chassistypes -eq 20){
            $chassisTypes = "Bus Expansion Chassis"
            }
            elseif($xml.info.win32_systemenclosure.chassistypes -eq 21){
            $chassisTypes = "Peripheral Chassis"
            }
            elseif($xml.info.win32_systemenclosure.chassistypes -eq 22){
            $chassisTypes = "Storage Chassis"
            }
            elseif($xml.info.win32_systemenclosure.chassistypes -eq 23){
            $chassisTypes = "Rack Mount Chassis"
            }
            elseif($xml.info.win32_systemenclosure.chassistypes -eq 24){
            $chassisTypes = "Sealed-Case PC"
            }
            else{
            $chassisTypes = "$chassisTypes [UNKNOWN]"
            }

        $table.cell(2,2).range.text = $chassisTypes
        $table.cell(3,1).range.text = "Lock Present"
        $table.cell(3,2).range.text = $xml.info.win32_systemenclosure.lockpresent
        $table.cell(4,1).range.text = "Manufacturer"
        $table.cell(4,2).range.text = $xml.info.win32_systemenclosure.manufacturer
        $table.cell(5,1).range.text = "Model"
        $table.cell(5,2).range.text = $xml.info.win32_systemenclosure.model
        $table.cell(6,1).range.text = "Security Status"

            if($xml.info.win32_systemenclosure.securitystatus -eq 1){
            $securityStatus = "Other"
            }
            elseif($xml.info.win32_systemenclosure.securitystatus -eq 2){
            $securityStatus = "Unknown"
            }
            elseif($xml.info.win32_systemenclosure.securitystatus -eq 3){
            $securityStatus = "None"
            }
            elseif($xml.info.win32_systemenclosure.securitystatus -eq 4){
            $securityStatus = "External Interface Locked Out"
            }
            elseif($xml.info.win32_systemenclosure.securitystatus -eq 5){
            $securityStatus = "External Interface Enabled"
            }
            else{
            $securityStatus = "$securityStatus [UNKNOWN]"
            }

        $table.cell(6,2).range.text = $securityStatus
        $table.cell(7,1).range.text = "Serial Number"
        $table.cell(7,2).range.text = $xml.info.win32_systemenclosure.serialnumber

        $table.Rows.item(1).Headingformat=-1
        $table.ApplyStyleFirstColumn = $false
        $selection.InsertCaption(-2, ": [WMI] System Enclosure", $null, 1, $false)

        }

    $selection.EndOf(15) | Out-Null
    $selection.MoveDown() | Out-Null
    $selection.InsertNewPage()
    #endregion ReportWMISystemEnclosure

    #region ReportWMIBootConfiguration
    # BASIC
    $selection.style = 'Heading 1'
    $selection.TypeText("[WMI] Boot Configuration")
    $selection.TypeParagraph()
    $selection.TypeParagraph()

    $selection.Style = 'Normal'
    $selection.TypeText("This data is collected from the Win32_BootConfiguration WMI class. For more information please go to https://docs.microsoft.com/en-gb/windows/desktop/CIMWin32Prov/win32-bootconfiguration")
    $selection.TypeParagraph()
    $selection.TypeParagraph()

        if(($xml.info.win32_bootconfiguration).count -eq 0){
        Write-Output "[WARNING] WMI Boot Configuration details not found. Moving on to next section..."
        $warnings ++
        $selection.Style = 'Normal'
        $selection.Font.Color="255"
        $selection.TypeText("[WARNING] WMI Boot Configuration data not found!")
        $selection.TypeParagraph()
        }
        elseif(($xml.info.win32_bootconfiguration.errorcode) -eq 1){
        Write-Output "[WARNING] WMI Boot Configuration details not collected. Moving on to next section..."
        $warnings ++
        $selection.Style = 'Normal'
        $selection.Font.Color="255"
        $selection.TypeText("[WARNING] WMI Boot Configuration data not collected!")
        $selection.TypeParagraph()
        $selection.TypeText("Reason for error: $($xml.info.win32_bootconfiguration.errortext)")
        $selection.TypeParagraph()
        }
        else{
        Write-Output "[INFO] Populating WMI Boot Configuration table."

        $table = $selection.Tables.add(
        $selection.Range,
        6,
        2,
        [Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior,
        [Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitWindow
        )

        $table.style = "Grid Table 4 - Accent 1"
        $table.cell(1,1).range.text = "Item"
        $table.cell(1,2).range.text = "Value"

        $table.cell(2,1).range.text = "Boot Directory"
        $table.cell(2,2).range.text = $xml.info.win32_bootconfiguration.bootdirectory
        $table.cell(3,1).range.text = "Caption"
        $table.cell(3,2).range.text = $xml.info.win32_bootconfiguration.caption
        $table.cell(4,1).range.text = "Configuration Path"
        $table.cell(4,2).range.text = $xml.info.win32_bootconfiguration.configurationpath
        $table.cell(5,1).range.text = "Description"
        $table.cell(5,2).range.text = $xml.info.win32_bootconfiguration.description
        $table.cell(6,1).range.text = "Scratch Directory"
        $table.cell(6,2).range.text = $xml.info.win32_bootconfiguration.scratchdirectory

        $table.Rows.item(1).Headingformat=-1
        $table.ApplyStyleFirstColumn = $false
        $selection.InsertCaption(-2, ": [WMI] Boot Configuration", $null, 1, $false)

        }

    $selection.EndOf(15) | Out-Null
    $selection.MoveDown() | Out-Null
    $selection.InsertNewPage()
    #endregion ReportWMIBootConfiguration

    #region ReportWMIBios
    # BASIC
    $selection.style = 'Heading 1'
    $selection.TypeText("[WMI] BIOS")
    $selection.TypeParagraph()
    $selection.TypeParagraph()

    $selection.Style = 'Normal'
    $selection.TypeText("This data is collected from the Win32_Bios WMI class. For more information please go to https://docs.microsoft.com/en-gb/windows/desktop/CIMWin32Prov/win32-bios")
    $selection.TypeParagraph()
    $selection.TypeParagraph()

        if(($xml.info.win32_bios).count -eq 0){
        Write-Output "[WARNING] WMI BIOS details not found. Moving on to next section..."
        $warnings ++
        $selection.Style = 'Normal'
        $selection.Font.Color="255"
        $selection.TypeText("[WARNING] WMI BIOS data not found!")
        $selection.TypeParagraph()
        }
        elseif(($xml.info.win32_bios.errorcode) -eq 1){
        Write-Output "[WARNING] WMI BIOS details not collected. Moving on to next section..."
        $warnings ++
        $selection.Style = 'Normal'
        $selection.Font.Color="255"
        $selection.TypeText("[WARNING] WMI BIOS data not collected!")
        $selection.TypeParagraph()
        $selection.TypeText("Reason for error: $($xml.info.win32_bios.errortext)")
        $selection.TypeParagraph()
        }
        else{
        Write-Output "[INFO] Populating WMI BIOS table."

        $table = $selection.Tables.add(
        $selection.Range,
        20,
        2,
        [Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior,
        [Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitWindow
        )

        $table.style = "Grid Table 4 - Accent 1"
        $table.cell(1,1).range.text = "Item"
        $table.cell(1,2).range.text = "Value"

        $table.cell(2,1).range.text = "BIOS Version"
        $table.cell(2,2).range.text = $xml.info.win32_bios.biosversion
        $table.cell(3,1).range.text = "BIOS Characteristics"

        $array = @(($xml.info.win32_bios.bioscharacteristics).split(" "))
        $bChars = @()

            foreach ($a in $array){

            $bChar = New-Object System.Object

                if($a -eq 0){
                $bChar | Add-Member -type NoteProperty -name v -Value "-_Reserved`r`n"
                $bChars += $bChar
                }
                elseif($a -eq 2){
                $bChar | Add-Member -type NoteProperty -name v -Value "-_Reserved`r`n"
                $bChars += $bChar
                }
                elseif($a -eq 3){
                $bChar | Add-Member -type NoteProperty -name v -Value "-_BIOS_Characteristics_not_supported`r`n"
                $bChars += $bChar
                }
                elseif($a -eq 4){
                $bChar | Add-Member -type NoteProperty -name v -Value "-_ISA_is_supported`r`n"
                $bChars += $bChar
                }
                elseif($a -eq 5){
                $bChar | Add-Member -type NoteProperty -name v -Value "-_MCA_is_supported`r`n"
                $bChars += $bChar
                }
                elseif($a -eq 6){
                $bChar | Add-Member -type NoteProperty -name v -Value "-_EISA_is_supported`r`n"
                $bChars += $bChar
                }
                elseif($a -eq 7){
                $bChar | Add-Member -type NoteProperty -name v -Value "-_PCI_is_supported`r`n"
                $bChars += $bChar
                }
                elseif($a -eq 8){
                $bChar | Add-Member -type NoteProperty -name v -Value "-_PC_Card_(PCMCIA)_is_supported`r`n"
                $bChars += $bChar
                }
                elseif($a -eq 9){
                $bChar | Add-Member -type NoteProperty -name v -Value "-_Plug_and_Play_is_supported`r`n"
                $bChars += $bChar
                }
                elseif($a -eq 10){
                $bChar | Add-Member -type NoteProperty -name v -Value "-_APM_is_supported`r`n"
                $bChars += $bChar
                }
                elseif($a -eq 11){
                $bChar | Add-Member -type NoteProperty -name v -Value "-_BIOS_is_upgradeable_(Flash)`r`n"
                $bChars += $bChar
                }
                elseif($a -eq 12){
                $bChar | Add-Member -type NoteProperty -name v -Value "-_BIOS_shadowing_is_allowed`r`n"
                $bChars += $bChar
                }
                elseif($a -eq 13){
                $bChar | Add-Member -type NoteProperty -name v -Value "-_VL-VESA_is_supported`r`n"
                $bChars += $bChar
                }
                elseif($a -eq 14){
                $bChar | Add-Member -type NoteProperty -name v -Value "-_ESCD_support_is_available`r`n"
                $bChars += $bChar
                }
                elseif($a -eq 15){
                $bChar | Add-Member -type NoteProperty -name v -Value "-_Boot_from_CD_is_supported`r`n"
                $bChars += $bChar
                }
                elseif($a -eq 16){
                $bChar | Add-Member -type NoteProperty -name v -Value "-_Selectable_Boot_is_supported`r`n"
                $bChars += $bChar
                }
                elseif($a -eq 17){
                $bChar | Add-Member -type NoteProperty -name v -Value "-_BIOS_ROM_is_socketed`r`n"
                $bChars += $bChar
                }
                elseif($a -eq 18){
                $bChar | Add-Member -type NoteProperty -name v -Value "-_Boot_from_PC_Card_(PCMCIA)_is_supported`r`n"
                $bChars += $bChar
                }
                elseif($a -eq 19){
                $bChar | Add-Member -type NoteProperty -name v -Value "-_EDD_(Enhanced_Disk_Drive)_Specification_is_supported`r`n"
                $bChars += $bChar
                }
                elseif($a -eq 20){
                $bChar | Add-Member -type NoteProperty -name v -Value "-_Int_13h_-_Japanese_Floppy_for_NEC_9800_1.2mb_(3.5\`",_1k_Bytes/Sector,_360_RPM)_is_supported`r`n"
                $bChars += $bChar
                }
                elseif($a -eq 21){
                $bChar | Add-Member -type NoteProperty -name v -Value "-_Int_13h_-_Japanese_Floppy_for_Toshiba_1.2mb_(3.5\`",_360_RPM)_is_supported`r`n"
                $bChars += $bChar
                }
                elseif($a -eq 22){
                $bChar | Add-Member -type NoteProperty -name v -Value "-_Int_13h_-_5.25\`"_/_360_KB_Floppy_Services_are_supported`r`n"
                $bChars += $bChar
                }
                elseif($a -eq 23){
                $bChar | Add-Member -type NoteProperty -name v -Value "-_Int_13h_-_5.25\`"_/1.2MB_Floppy_Services_are_supported`r`n"
                $bChars += $bChar
                }
                elseif($a -eq 24){
                $bChar | Add-Member -type NoteProperty -name v -Value "-_Int_13h_-_3.5\`"_/_720_KB_Floppy_Services_are_supported`r`n"
                $bChars += $bChar
                }
                elseif($a -eq 25){
                $bChar | Add-Member -type NoteProperty -name v -Value "-_Int_13h_-_3.5\`"_/_2.88_MB_Floppy_Services_are_supported`r`n"
                $bChars += $bChar
                }
                elseif($a -eq 26){
                $bChar | Add-Member -type NoteProperty -name v -Value "-_Int_5h,_Print_Screen_Service_is_supported`r`n"
                $bChars += $bChar
                }
                elseif($a -eq 27){
                $bChar | Add-Member -type NoteProperty -name v -Value "-_Int_9h,_8042_Keyboard_Services_are_supported`r`n"
                $bChars += $bChar
                }
                elseif($a -eq 28){
                $bChar | Add-Member -type NoteProperty -name v -Value "-_Int_14h,_Serial_Services_are_supported`r`n"
                $bChars += $bChar
                }
                elseif($a -eq 29){
                $bChar | Add-Member -type NoteProperty -name v -Value "-_Int_17h,_Printer_Services_are_supported`r`n"
                $bChars += $bChar
                }
                elseif($a -eq 30){
                $bChar | Add-Member -type NoteProperty -name v -Value "-_Int_10h,_CGA/Mono_Video_Services_are_supported`r`n"
                $bChars += $bChar
                }
                elseif($a -eq 31){
                $bChar | Add-Member -type NoteProperty -name v -Value "-_NEC_PC-98`r`n"
                $bChars += $bChar
                }
                elseif($a -eq 32){
                $bChar | Add-Member -type NoteProperty -name v -Value "-_ACPI_supported`r`n"
                $bChars += $bChar
                }
                elseif($a -eq 33){
                $bChar | Add-Member -type NoteProperty -name v -Value "-_USB_Legacy_is_supported`r`n"
                $bChars += $bChar
                }
                elseif($a -eq 34){
                $bChar | Add-Member -type NoteProperty -name v -Value "-_AGP_is_supported`r`n"
                $bChars += $bChar
                }
                elseif($a -eq 35){
                $bChar | Add-Member -type NoteProperty -name v -Value "-_I2O_Boot_is_supported`r`n"
                $bChars += $bChar
                }
                elseif($a -eq 36){
                $bChar | Add-Member -type NoteProperty -name v -Value "-_LS-120_Boot_is_supported`r`n"
                $bChars += $bChar
                }
                elseif($a -eq 37){
                $bChar | Add-Member -type NoteProperty -name v -Value "-_ATAPI_ZIP_Drive_Boot_is_supported`r`n"
                $bChars += $bChar
                }
                elseif($a -eq 38){
                $bChar | Add-Member -type NoteProperty -name v -Value "-_1394_Boot_is_supported`r`n"
                $bChars += $bChar
                }
                elseif($a -eq 39){
                $bChar | Add-Member -type NoteProperty -name v -Value "-_Smart_Battery_supported`r`n"
                $bChars += $bChar
                }
                elseif(($a -ge 40) -and ($a -le 47)){
                $bChar | Add-Member -type NoteProperty -name v -Value "-_$a`_Reserved_for_BIOS_vendor`r`n"
                $bChars += $bChar
                }
                elseif(($a -ge 48) -and ($a -le 63)){
                $bChar | Add-Member -type NoteProperty -name v -Value "-_$a`_Reserved_for_system_vendor`r`n"
                $bChars += $bChar
                }
                else{
                $bChar | Add-Member -type NoteProperty -name v -Value "-_$a`_[UNKNOWN]`r`n"
                $bChars += $bChar
                }

        }

        # Array entries had space at start after 1st entry, for example:
        # Entry1
        #  Entry 2
        #  Entry 3
        # To sort, add _ to text above, use replace to remove space then use replace to change _ to space.
        $bC = ([string]$bChars.v -replace " ", "") -replace "_", " "

        $table.cell(3,2).range.text = "$bC"
        $table.cell(4,1).range.text = "Build Number"
        $table.cell(4,2).range.text = $xml.info.win32_bios.buildnumber
        $table.cell(5,1).range.text = "Caption"
        $table.cell(5,2).range.text = $xml.info.win32_bios.caption
        $table.cell(6,1).range.text = "Description"
        $table.cell(6,2).range.text = $xml.info.win32_bios.description
        $table.cell(7,1).range.text = "Manufacturer"
        $table.cell(7,2).range.text = $xml.info.win32_bios.manufacturer
        $table.cell(8,1).range.text = "Name"
        $table.cell(8,2).range.text = $xml.info.win32_bios.name
        $table.cell(9,1).range.text = "Primary BIOS"
        $table.cell(9,2).range.text = $xml.info.win32_bios.primarybios
        $table.cell(10,1).range.text = "Release Date"

        $relDate =
        $xml.info.win32_bios.releasedate.Substring(6,2) + "/" +
        $xml.info.win32_bios.releasedate.Substring(4,2) + "/" +
        $xml.info.win32_bios.releasedate.Substring(0,4)

        $table.cell(10,2).range.text = $relDate
        $table.cell(11,1).range.text = "SM BIOS Version"
        $table.cell(11,2).range.text = $xml.info.win32_bios.smbiosbiosversion
        $table.cell(12,1).range.text = "SM BIOS Major Version"
        $table.cell(12,2).range.text = $xml.info.win32_bios.smbiosmajorversion
        $table.cell(13,1).range.text = "SM BIOS Minor Version"
        $table.cell(13,2).range.text = $xml.info.win32_bios.smbiosminorversion
        $table.cell(14,1).range.text = "SM BIOS Present"
        $table.cell(14,2).range.text = $xml.info.win32_bios.smbiospresent
        $table.cell(15,1).range.text = "Serial Number"
        $table.cell(15,2).range.text = $xml.info.win32_bios.serialnumber
        $table.cell(16,1).range.text = "Software Element ID"
        $table.cell(16,2).range.text = $xml.info.win32_bios.softwareelementid
        $table.cell(17,1).range.text = "Software Element State"

            if($xml.info.win32_bios.softwareelementstate -eq 0){
            $sES = "Deployable"
            }
		    elseif($xml.info.win32_bios.softwareelementstate -eq 1){
            $sES = "Installable"
            }
            elseif($xml.info.win32_bios.softwareelementstate -eq 2){
            $sES = "Executable"
            }
            elseif($xml.info.win32_bios.softwareelementstate -eq 3){
            $sES = "Running"
            }
            else{
            $sES = "$sES [UNKNOWN]"
            }

        $table.cell(17,2).range.text = $sES
        $table.cell(18,1).range.text = "System BIOS Major Version"
        $table.cell(18,2).range.text = $xml.info.win32_bios.systembiosmajorversion
        $table.cell(19,1).range.text = "System BIOS Minor Version"
        $table.cell(19,2).range.text = $xml.info.win32_bios.systembiosminorversion
        $table.cell(20,1).range.text = "Version"
        $table.cell(20,2).range.text = $xml.info.win32_bios.version

        $table.Rows.item(1).Headingformat=-1
        $table.ApplyStyleFirstColumn = $false
        $selection.InsertCaption(-2, ": [WMI] BIOS", $null, 1, $false)

        }

    $selection.EndOf(15) | Out-Null
    $selection.MoveDown() | Out-Null
    $selection.InsertNewPage()
    #endregion ReportWMIBios

    #region ReportWMIUserAccounts
    # BASIC & DETAILED
    $selection.style = 'Heading 1'
    $selection.TypeText("[WMI] User Accounts")
    $selection.TypeParagraph()
    $selection.TypeParagraph()

    $selection.Style = 'Normal'
    $selection.TypeText("This data is collected from the Win32_UserAccounts WMI class. For more information please go to https://docs.microsoft.com/en-gb/windows/desktop/CIMWin32Prov/win32-useraccount")
    $selection.TypeParagraph()
    $selection.TypeParagraph()

        if($reportType -ne 'Basic'){
        $selection.style = 'Heading 2'
        $selection.TypeText("Basic")
        $selection.TypeParagraph()
        $selection.TypeParagraph()
        }

        if(($xml.info.win32_useraccount).count -eq 0){
        Write-Output "[WARNING] WMI User Accounts (Basic) details not found. Moving on to next section..."
        $warnings ++
        $selection.Style = 'Normal'
        $selection.Font.Color="255"
        $selection.TypeText("[WARNING] WMI User Accounts (Basic) data not found!")
        $selection.TypeParagraph()
        }
        elseif(($xml.info.win32_useraccount.errorcode) -eq 1){
        Write-Output "[WARNING] WMI User Accounts (Basic) details not collected. Moving on to next section..."
        $warnings ++
        $selection.Style = 'Normal'
        $selection.Font.Color="255"
        $selection.TypeText("[WARNING] WMI User Accounts (Basic) data not collected!")
        $selection.TypeParagraph()
        $selection.TypeText("Reason for error: $($xml.info.win32_useraccount.errortext)")
        $selection.TypeParagraph()
        }
        else{
        Write-Output "[INFO] Populating WMI User (Basic) Accounts table."

        # Count rows in multi
        $multis = $xml.selectnodes("//info/win32_useraccount/win32_useraccount_multi")
        $count = ($multis | Measure-Object).Count

        # Calculate rows (from query above) and add 1 for the header
        $rows = $count + 1

        $table = $selection.Tables.add(
        $selection.Range,
        $rows,
        6,
        [Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior,
        [Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitWindow
        )

        $table.style = "Grid Table 4 - Accent 1"
        $table.cell(1,1).range.text = "Domain"
        $table.cell(1,2).range.text = "Name"
        $table.cell(1,3).range.text = "Disabled"
        $table.cell(1,4).range.text = "Password Expires"
        $table.cell(1,5).range.text = "Password Required"
        $table.cell(1,6).range.text = "Locked Out"

        $i = 2 # 2 as header is row 1

            foreach($multi in $multis){

            $table.cell($i,1).range.text = $multi.domain
            $table.cell($i,2).range.text = $multi.name
            $table.cell($i,3).range.text = $multi.disabled
            $table.cell($i,4).range.text = $multi.passwordexpires
            $table.cell($i,5).range.text = $multi.passwordrequired
            $table.cell($i,6).range.text = $multi.lockout
            $i++

            }

        $table.Rows.item(1).Headingformat=-1
        $table.ApplyStyleFirstColumn = $false
        $selection.InsertCaption(-2, ": [WMI] User Account (Basic)", $null, 1, $false)

          if($reportType -ne "Basic"){
          $selection.TypeParagraph()
          }

        }

    $selection.EndOf(15) | Out-Null
    $selection.MoveDown() | Out-Null

        if($reportType -eq "Detailed"){
        $selection.TypeParagraph()
        $selection.style = 'Heading 2'
        $selection.TypeText("Detailed")
        $selection.TypeParagraph()
        $selection.TypeParagraph()

            if(($xml.info.win32_useraccount).count -eq 0){
            Write-Output "[WARNING] WMI User Accounts (Detailed) details not found. Moving on to next section..."
            $warnings ++
            $selection.Style = 'Normal'
            $selection.Font.Color="255"
            $selection.TypeText("[WARNING] WMI User Accounts (Detailed) data not found!")
            $selection.TypeParagraph()
            }
            elseif(($xml.info.win32_useraccount.errorcode) -eq 1){
            Write-Output "[WARNING] WMI User Accounts (Detailed) details not collected. Moving on to next section..."
            $warnings ++
            $selection.Style = 'Normal'
            $selection.Font.Color="255"
            $selection.TypeText("[WARNING] WMI User Accounts (Detailed) data not collected!")
            $selection.TypeParagraph()
            $selection.TypeText("Reason for error: $($xml.info.win32_useraccount.errortext)")
            $selection.TypeParagraph()
            }
            else{
            Write-Output "[INFO] Populating WMI User (Detailed) Accounts table."

            # Count rows in multi
            $multis = $xml.selectnodes("//info/win32_useraccount/win32_useraccount_multi")
            $count = ($multis | Measure-Object).Count

                # Calculate rows. (No of items x no of multis) + (no of multis + 1 (for header)) | or less 1 for header if single row
                if($count -eq 1){
                $rows = (14 + $count)
                }
                else{
                $rows = (14 * $count) + ($count + 1)
                }

            $table = $selection.Tables.add(
            $selection.Range,
            $rows,
            2,
            [Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior,
            [Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitWindow
            )

            $table.style = "Grid Table 4 - Accent 1"
            $table.cell(1,1).range.text = "Item"
            $table.cell(1,2).range.text = "Value"

            $i = 2 # 2 as header is row 1
            $y = 1

                foreach($multi in $multis){

                    if($count -gt 1){
                    $table.cell($i,1).Merge($table.cell($i, 2))
                    $table.cell($i,1).range.text = "------ BLOCK $y ------"
                    $i++
                    }

                $table.cell($i,1).range.text = "Account Type"

        	        if($multi.accounttype -eq 256){
                    $accountType = "Temporary duplicate account"
                    }
                    elseif($multi.accounttype -eq 512){
                    $accountType = "Normal account"
                    }
                    elseif($multi.accounttype -eq 2048){
                    $accountType = "Interdomain trust account"
                    }
                    elseif($multi.accounttype -eq 4096){
                    $accountType = "Workstation trust account"
                    }
                    elseif($multi.accounttype -eq 8192){
                    $accountType = "Server trust account"
                    }
		            else{
		            $accountType = "$multi.accounttype [UNKNOWN]"
		            }

                $table.cell($i,2).range.text = $accountType
                $i++

                $table.cell($i,1).range.text = "Caption"
                $table.cell($i,2).range.text = $multi.caption
                $i++

                $table.cell($i,1).range.text = "Description"
                $table.cell($i,2).range.text = $multi.description
                $i++

                $table.cell($i,1).range.text = "Disabled"
                $table.cell($i,2).range.text = $multi.disabled
                $i++

                $table.cell($i,1).range.text = "Domain"
                $table.cell($i,2).range.text = $multi.domain
                $i++

                $table.cell($i,1).range.text = "Fullname"
                $table.cell($i,2).range.text = $multi.fullname
                $i++

                $table.cell($i,1).range.text = "Local Account"
                $table.cell($i,2).range.text = $multi.localaccount
                $i++

                $table.cell($i,1).range.text = "Locked Out"
                $table.cell($i,2).range.text = $multi.lockout
                $i++

                $table.cell($i,1).range.text = "Name"
                $table.cell($i,2).range.text = $multi.name
                $i++

                $table.cell($i,1).range.text = "Password Changeable"
                $table.cell($i,2).range.text = $multi.passwordchangeable
                $i++

                $table.cell($i,1).range.text = "Password Expires"
                $table.cell($i,2).range.text = $multi.passwordexpires
                $i++

                $table.cell($i,1).range.text = "Password Required"
                $table.cell($i,2).range.text = $multi.passwordrequired
                $i++

                $table.cell($i,1).range.text = "SID"
                $table.cell($i,2).range.text = $multi.sid
                $i++

                $table.cell($i,1).range.text = "SID Type"

        	        if($multi.sidtype -eq 1){
                    $sIDType = "User"
                    }
                    elseif($multi.sidtype -eq 2){
                    $sIDType = "Group"
                    }
                    elseif($multi.sidtype -eq 3){
                    $sIDType = "Domain"
                    }
                    elseif($multi.sidtype -eq 4){
                    $sIDType = "Alias"
                    }
                    elseif($multi.sidtype -eq 5){
                    $sIDType = "Well Known Group"
                    }
                    elseif($multi.sidtype -eq 6){
                    $sIDType = "Deleted Account"
                    }
                    elseif($multi.sidtype -eq 7){
                    $sIDType = "Invalid"
                    }
                    elseif($multi.sidtype -eq 8){
                    $sIDType = "Unknown"
                    }
                    elseif($multi.sidtype -eq 9){
                    $sIDType = "Computer"
                    }
		            else{
		            $sIDType = "$multi.sidtype [UNKNOWN]"
		            }

                $table.cell($i,2).range.text = $sIDType
                $i++

                $y++

                }

            $table.Rows.item(1).Headingformat=-1
            $table.ApplyStyleFirstColumn = $false
            $selection.InsertCaption(-2, ": [WMI] User Account (Detailed)", $null, 1, $false)

            }

        }

    $selection.EndOf(15) | Out-Null
    $selection.MoveDown() | Out-Null
    $selection.InsertNewPage()
    #endregion ReportWMIUserAccounts

    #region ReportWMIGroups
    # BASIC & DETAILED
    $selection.style = 'Heading 1'
    $selection.TypeText("[WMI] Groups")
    $selection.TypeParagraph()
    $selection.TypeParagraph()

    $selection.Style = 'Normal'
    $selection.TypeText("This data is collected from the Win32_Group WMI class. For more information please go to https://docs.microsoft.com/en-gb/windows/desktop/CIMWin32Prov/win32-group")
    $selection.TypeParagraph()
    $selection.TypeParagraph()

        if($reportType -ne 'Basic'){
        $selection.style = 'Heading 2'
        $selection.TypeText("Basic")
        $selection.TypeParagraph()
        $selection.TypeParagraph()
        }

        if(($xml.info.win32_group).count -eq 0){
        Write-Output "[WARNING] WMI Group (Basic) details not found. Moving on to next section..."
        $warnings ++
        $selection.Style = 'Normal'
        $selection.Font.Color="255"
        $selection.TypeText("[WARNING] WMI Group (Basic) data not found!")
        $selection.TypeParagraph()
        }
        elseif(($xml.info.win32_group.errorcode) -eq 1){
        Write-Output "[WARNING] WMI Group (Basic) details not collected. Moving on to next section..."
        $warnings ++
        $selection.Style = 'Normal'
        $selection.Font.Color="255"
        $selection.TypeText("[WARNING] WMI Group (Basic) data not collected!")
        $selection.TypeParagraph()
        $selection.TypeText("Reason for error: $($xml.info.win32_group.errortext)")
        $selection.TypeParagraph()
        }
        else{
        Write-Output "[INFO] Populating WMI Group (Basic) table."

        # Count rows in multi
        $multis = $xml.selectnodes("//info/win32_group/win32_group_multi")
        $count = ($multis | Measure-Object).Count

        # Calculate rows (from query above) and add 1 for the header
        $rows = $count + 1

        $table = $selection.Tables.add(
        $selection.Range,
        $rows,
        3,
        [Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior,
        [Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitWindow
        )

        $table.style = "Grid Table 4 - Accent 1"
        $table.cell(1,1).range.text = "Domain"
        $table.cell(1,2).range.text = "Name"
        $table.cell(1,3).range.text = "Local Group"

        $i = 2 # 2 as header is row 1

            foreach($multi in $multis){

            $table.cell($i,1).range.text = $multi.domain
            $table.cell($i,2).range.text = $multi.name
            $table.cell($i,3).range.text = $multi.localaccount
            $i++
            }

        $table.Rows.item(1).Headingformat=-1
        $table.ApplyStyleFirstColumn = $false
        $selection.InsertCaption(-2, ": [WMI] Group (Basic)", $null, 1, $false)

          if($reportType -ne "Basic"){
          $selection.TypeParagraph()
          }

        }

    $selection.EndOf(15) | Out-Null
    $selection.MoveDown() | Out-Null

        if($reportType -eq "Detailed"){
        $selection.TypeParagraph()
        $selection.style = 'Heading 2'
        $selection.TypeText("Detailed")
        $selection.TypeParagraph()
        $selection.TypeParagraph()

            if(($xml.info.win32_group).count -eq 0){
            Write-Output "[WARNING] WMI Group (Detailed) details not found. Moving on to next section..."
            $warnings ++
            $selection.Style = 'Normal'
            $selection.Font.Color="255"
            $selection.TypeText("[WARNING] WMI Group (Detailed) data not found!")
            $selection.TypeParagraph()
            }
            elseif(($xml.info.win32_group.errorcode) -eq 1){
            Write-Output "[WARNING] WMI Group (Detailed) details not collected. Moving on to next section..."
            $warnings ++
            $selection.Style = 'Normal'
            $selection.Font.Color="255"
            $selection.TypeText("[WARNING] WMI Group (Detailed) data not collected!")
            $selection.TypeParagraph()
            $selection.TypeText("Reason for error: $($xml.info.win32_group.errortext)")
            $selection.TypeParagraph()
            }
            else{
            Write-Output "[INFO] Populating WMI Group (Detailed) table."

            # Count rows in multi
            $multis = $xml.selectnodes("//info/win32_group/win32_group_multi")
            $count = ($multis | Measure-Object).Count

                # Calculate rows. (No of items x no of multis) + (no of multis + 1 (for header)) | or less 1 for header if single row
                if($count -eq 1){
                $rows = (7 + $count)
                }
                else{
                $rows = (7 * $count) + ($count + 1)
                }

            $table = $selection.Tables.add(
            $selection.Range,
            $rows,
            2,
            [Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior,
            [Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitWindow
            )

            $table.style = "Grid Table 4 - Accent 1"
            $table.cell(1,1).range.text = "Item"
            $table.cell(1,2).range.text = "Value"

            $i = 2 # 2 as header is row 1
            $y = 1

                foreach($multi in $multis){

                    if($count -gt 1){
                    $table.cell($i,1).Merge($table.cell($i, 2))
                    $table.cell($i,1).range.text = "------ BLOCK $y ------"
                    $i++
                    }

                $table.cell($i,1).range.text = "Caption"
                $table.cell($i,2).range.text = $multi.caption
                $i++

                $table.cell($i,1).range.text = "Description"
                $table.cell($i,2).range.text = $multi.description
                $i++

                $table.cell($i,1).range.text = "Domain"
                $table.cell($i,2).range.text = $multi.domain
                $i++

                $table.cell($i,1).range.text = "Local Group"
                $table.cell($i,2).range.text = $multi.localaccount
                $i++

                $table.cell($i,1).range.text = "Name"
                $table.cell($i,2).range.text = $multi.name
                $i++

                $table.cell($i,1).range.text = "SID"
                $table.cell($i,2).range.text = $multi.sid
                $i++

                $table.cell($i,1).range.text = "SID Type"

        	        if($multi.sidtype -eq 1){
                    $sIDType = "User"
                    }
                    elseif($multi.sidtype -eq 2){
                    $sIDType = "Group"
                    }
                    elseif($multi.sidtype -eq 3){
                    $sIDType = "Domain"
                    }
                    elseif($multi.sidtype -eq 4){
                    $sIDType = "Alias"
                    }
                    elseif($multi.sidtype -eq 5){
                    $sIDType = "Well Known Group"
                    }
                    elseif($multi.sidtype -eq 6){
                    $sIDType = "Deleted Account"
                    }
                    elseif($multi.sidtype -eq 7){
                    $sIDType = "Invalid"
                    }
                    elseif($multi.sidtype -eq 8){
                    $sIDType = "Unknown"
                    }
                    elseif($multi.sidtype -eq 9){
                    $sIDType = "Computer"
                    }
		            else{
		            $sIDType = "$multi.sidtype [UNKNOWN]"
		            }

                $table.cell($i,2).range.text = $sIDType
                $i++

                $y++

                }

            $table.Rows.item(1).Headingformat=-1
            $table.ApplyStyleFirstColumn = $false
            $selection.InsertCaption(-2, ": [WMI] Group (Detailed)", $null, 1, $false)

            }

        }

    $selection.EndOf(15) | Out-Null
    $selection.MoveDown() | Out-Null
    $selection.InsertNewPage()
    #endregion ReportWMIGroups

    #region ReportWMIGroupMembership
    # BASIC
    $selection.style = 'Heading 1'
    $selection.TypeText("[WMI] Group Membership")
    $selection.TypeParagraph()
    $selection.TypeParagraph()

    $selection.Style = 'Normal'
    $selection.TypeText("This data is collected from the Win32_Group & Win32_GroupUser WMI class. For more information please go to https://docs.microsoft.com/en-gb/windows/desktop/CIMWin32Prov/win32-group & https://docs.microsoft.com/en-gb/windows/desktop/CIMWin32Prov/win32-groupuser")
    $selection.TypeParagraph()
    $selection.TypeParagraph()

        if(($xml.info.win32_group_membership).count -eq 0){
        Write-Output "[WARNING] WMI Group Membership details not found. Moving on to next section..."
        $warnings ++
        $selection.Style = 'Normal'
        $selection.Font.Color="255"
        $selection.TypeText("[WARNING] WMI Group Membership data not found!")
        $selection.TypeParagraph()
        }
        elseif(($xml.info.win32_group_membership.errorcode) -eq 1){
        Write-Output "[WARNING] WMI Group Membership details not collected. Moving on to next section..."
        $warnings ++
        $selection.Style = 'Normal'
        $selection.Font.Color="255"
        $selection.TypeText("[WARNING] WMI Group Membership data not collected!")
        $selection.TypeParagraph()
        $selection.TypeText("Reason for error: $($xml.info.win32_group_membership.errortext)")
        $selection.TypeParagraph()
        }
        else{
        Write-Output "[INFO] Populating WMI Group Membership table."

        # Count rows in multi
        $multis = $xml.selectnodes("//info/win32_group_membership/win32_group_multi")
        $count = ($multis | Measure-Object).Count


        $z = 0
        foreach($multi in $multis){

            # Count rows with members, dont get blank rows
            if(([string]$multi.members).length -gt 0){
            $z++
            }

        }

        # Calculate rows (from query above) and add 1 for the header
        $rows = $z + 1

        $table = $selection.Tables.add(
        $selection.Range,
        $rows,
        2,
        [Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior,
        [Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitWindow
        )

        $table.style = "Grid Table 4 - Accent 1"
        $table.cell(1,1).range.text = "Group"
        $table.cell(1,2).range.text = "Members"

        $i = 2 # 2 as header is row 1

            foreach($multi in $multis){

                # dont get rows with blank members
                if(([string]$multi.members).length -ne 0){

                $table.cell($i,1).range.text = $multi.name

                # Join members together using @ then replace @ with linefeed
                $mems =  (($multi.members.member) -join "@") -replace "@", "`r`n"

                $table.cell($i,2).range.text = "$mems"
                $i++

                }

            }

        $table.Rows.item(1).Headingformat=-1
        $table.ApplyStyleFirstColumn = $false
        $selection.InsertCaption(-2, ": [WMI] Group Membership", $null, 1, $false)

        }

    $selection.EndOf(15) | Out-Null
    $selection.MoveDown() | Out-Null
    $selection.InsertNewPage()
    #endregion ReportWMIGroupMembership

    #region ReportWMISystemAccounts
    # BASIC & DETAILED
    $selection.style = 'Heading 1'
    $selection.TypeText("[WMI] System Accounts")
    $selection.TypeParagraph()
    $selection.TypeParagraph()

    $selection.Style = 'Normal'
    $selection.TypeText("This data is collected from the Win32_SystemAccount WMI class. For more information please go to https://docs.microsoft.com/en-gb/windows/desktop/CIMWin32Prov/win32-systemaccount")
    $selection.TypeParagraph()
    $selection.TypeParagraph()

        if($reportType -ne 'Basic'){
        $selection.style = 'Heading 2'
        $selection.TypeText("Basic")
        $selection.TypeParagraph()
        $selection.TypeParagraph()
        }

        if(($xml.info.win32_systemaccount).count -eq 0){
        Write-Output "[WARNING] WMI System Accounts (Basic) details not found. Moving on to next section..."
        $warnings ++
        $selection.Style = 'Normal'
        $selection.Font.Color="255"
        $selection.TypeText("[WARNING] WMI System Accounts (Basic) data not found!")
        $selection.TypeParagraph()
        }
        elseif(($xml.info.win32_systemaccount.errorcode) -eq 1){
        Write-Output "[WARNING] WMI System Accounts (Basic) details not collected. Moving on to next section..."
        $warnings ++
        $selection.Style = 'Normal'
        $selection.Font.Color="255"
        $selection.TypeText("[WARNING] WMI System Accounts (Basic) data not collected!")
        $selection.TypeParagraph()
        $selection.TypeText("Reason for error: $($xml.info.win32_systemaccount.errortext)")
        $selection.TypeParagraph()
        }
        else{
        Write-Output "[INFO] Populating WMI System Accounts (Basic) table."

        # Count rows in multi
        $multis = $xml.selectnodes("//info/win32_systemaccount/win32_systemaccount_multi")
        $count = ($multis | Measure-Object).Count

        # Calculate rows (from query above) and add 1 for the header
        $rows = $count + 1

        $table = $selection.Tables.add(
        $selection.Range,
        $rows,
        3,
        [Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior,
        [Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitWindow
        )

        $table.style = "Grid Table 4 - Accent 1"
        $table.cell(1,1).range.text = "Domain"
        $table.cell(1,2).range.text = "Name"
        $table.cell(1,3).range.text = "Local Account"

        $i = 2 # 2 as header is row 1

            foreach($multi in $multis){

            $table.cell($i,1).range.text = $multi.domain
            $table.cell($i,2).range.text = $multi.name
            $table.cell($i,3).range.text = $multi.localaccount
            $i++

            }

        $table.Rows.item(1).Headingformat=-1
        $table.ApplyStyleFirstColumn = $false
        $selection.InsertCaption(-2, ": [WMI] System Account (Basic)", $null, 1, $false)

          if($reportType -ne "Basic"){
          $selection.TypeParagraph()
          }

        }

    $selection.EndOf(15) | Out-Null
    $selection.MoveDown() | Out-Null

        if($reportType -eq "Detailed"){
        $selection.TypeParagraph()
        $selection.style = 'Heading 2'
        $selection.TypeText("Detailed")
        $selection.TypeParagraph()
        $selection.TypeParagraph()

            if(($xml.info.win32_systemaccount).count -eq 0){
            Write-Output "[WARNING] WMI System Account (Detailed) details not found. Moving on to next section..."
            $warnings ++
            $selection.Style = 'Normal'
            $selection.Font.Color="255"
            $selection.TypeText("[WARNING] WMI System Account (Detailed) data not found!")
            $selection.TypeParagraph()
            }
            elseif(($xml.info.win32_systemaccount.errorcode) -eq 1){
            Write-Output "[WARNING] WMI System Accounts (Detailed) details not collected. Moving on to next section..."
            $warnings ++
            $selection.Style = 'Normal'
            $selection.Font.Color="255"
            $selection.TypeText("[WARNING] WMI System Accounts (Detailed) data not collected!")
            $selection.TypeParagraph()
            $selection.TypeText("Reason for error: $($xml.info.win32_systemaccount.errortext)")
            $selection.TypeParagraph()
            }
            else{
            Write-Output "[INFO] Populating WMI System Account (Detailed) table."

            # Count rows in multi
            $multis = $xml.selectnodes("//info/win32_systemaccount/win32_systemaccount_multi")
            $count = ($multis | Measure-Object).Count

                # Calculate rows. (No of items x no of multis) + (no of multis + 1 (for header)) | or less 1 for header if single row
                if($count -eq 1){
                $rows = (7 + $count)
                }
                else{
                $rows = (7 * $count) + ($count + 1)
                }

            $table = $selection.Tables.add(
            $selection.Range,
            $rows,
            2,
            [Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior,
            [Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitWindow
            )

            $table.style = "Grid Table 4 - Accent 1"
            $table.cell(1,1).range.text = "Item"
            $table.cell(1,2).range.text = "Value"

            $i = 2 # 2 as header is row 1
            $y = 1

                foreach($multi in $multis){

                    if($count -gt 1){
                    $table.cell($i,1).Merge($table.cell($i, 2))
                    $table.cell($i,1).range.text = "------ BLOCK $y ------"
                    $i++
                    }

                $table.cell($i,1).range.text = "Caption"
                $table.cell($i,2).range.text = $multi.caption
                $i++

                $table.cell($i,1).range.text = "Description"
                $table.cell($i,2).range.text = $multi.description
                $i++

                $table.cell($i,1).range.text = "Domain"
                $table.cell($i,2).range.text = $multi.domain
                $i++

                $table.cell($i,1).range.text = "Local Account"
                $table.cell($i,2).range.text = $multi.localaccount
                $i++

                $table.cell($i,1).range.text = "Name"
                $table.cell($i,2).range.text = $multi.name
                $i++

                $table.cell($i,1).range.text = "SID"
                $table.cell($i,2).range.text = $multi.sid
                $i++

                $table.cell($i,1).range.text = "SID Type"

        	        if($multi.sidtype -eq 1){
                    $sIDType = "User"
                    }
                    elseif($multi.sidtype -eq 2){
                    $sIDType = "Group"
                    }
                    elseif($multi.sidtype -eq 3){
                    $sIDType = "Domain"
                    }
                    elseif($multi.sidtype -eq 4){
                    $sIDType = "Alias"
                    }
                    elseif($multi.sidtype -eq 5){
                    $sIDType = "Well Known Group"
                    }
                    elseif($multi.sidtype -eq 6){
                    $sIDType = "Deleted Account"
                    }
                    elseif($multi.sidtype -eq 7){
                    $sIDType = "Invalid"
                    }
                    elseif($multi.sidtype -eq 8){
                    $sIDType = "Unknown"
                    }
                    elseif($multi.sidtype -eq 9){
                    $sIDType = "Computer"
                    }
		            else{
		            $sIDType = "$multi.sidtype [UNKNOWN]"
		            }

                $table.cell($i,2).range.text = $sIDType
                $i++

                $y++

                }

            $table.Rows.item(1).Headingformat=-1
            $table.ApplyStyleFirstColumn = $false
            $selection.InsertCaption(-2, ": [WMI] System Account (Detailed)", $null, 1, $false)

            }

        }

    $selection.EndOf(15) | Out-Null
    $selection.MoveDown() | Out-Null
    $selection.InsertNewPage()
    #endregion ReportWMISystemAccounts

    #region ReportWMITimezone
    # BASIC
    $selection.style = 'Heading 1'
    $selection.TypeText("[WMI] Time Zone")
    $selection.TypeParagraph()
    $selection.TypeParagraph()

    $selection.Style = 'Normal'
    $selection.TypeText("This data is collected from the Win32_TimeZone WMI class. For more information please go to https://docs.microsoft.com/en-gb/windows/desktop/CIMWin32Prov/win32-timezone")
    $selection.TypeParagraph()
    $selection.TypeParagraph()

        if(($xml.info.win32_timezone).count -eq 0){
        Write-Output "[WARNING] WMI Time Zone details not found. Moving on to next section..."
        $warnings ++
        $selection.Style = 'Normal'
        $selection.Font.Color="255"
        $selection.TypeText("[WARNING] WMI Time Zone data not found!")
        $selection.TypeParagraph()
        }
        elseif(($xml.info.win32_timezone.errorcode) -eq 1){
        Write-Output "[WARNING] WMI Timezone details not collected. Moving on to next section..."
        $warnings ++
        $selection.Style = 'Normal'
        $selection.Font.Color="255"
        $selection.TypeText("[WARNING] WMI Timezone data not collected!")
        $selection.TypeParagraph()
        $selection.TypeText("Reason for error: $($xml.info.win32_timezone.errortext)")
        $selection.TypeParagraph()
        }
        else{
        Write-Output "[INFO] Populating WMI Time Zone table."

        $table = $selection.Tables.add(
        $selection.Range,
        3,
        2,
        [Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior,
        [Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitWindow
        )

        $table.style = "Grid Table 4 - Accent 1"
        $table.cell(1,1).range.text = "Item"
        $table.cell(1,2).range.text = "Value"

        $table.cell(2,1).range.text = "Caption"
        $table.cell(2,2).range.text = $xml.info.win32_timezone.caption
        $table.cell(3,1).range.text = "Standard Name"
        $table.cell(3,2).range.text = $xml.info.win32_timezone.standardname

        $table.Rows.item(1).Headingformat=-1
        $table.ApplyStyleFirstColumn = $false
        $selection.InsertCaption(-2, ": [WMI] Time Zone", $null, 1, $false)

        }

    $selection.EndOf(15) | Out-Null
    $selection.MoveDown() | Out-Null
    $selection.InsertNewPage()
    #endregion ReportWMITimezone

    #region ReportWMIRegistry
    # BASIC
    $selection.style = 'Heading 1'
    $selection.TypeText("[WMI] Registry")
    $selection.TypeParagraph()
    $selection.TypeParagraph()

    $selection.Style = 'Normal'
    $selection.TypeText("This data is collected from the Win32_Registry WMI class. For more information please go to https://docs.microsoft.com/en-gb/windows/desktop/CIMWin32Prov/win32-registry")
    $selection.TypeParagraph()
    $selection.TypeParagraph()

        if(($xml.info.win32_registry).count -eq 0){
        Write-Output "[WARNING] WMI Registry details not found. Moving on to next section..."
        $warnings ++
        $selection.Style = 'Normal'
        $selection.Font.Color="255"
        $selection.TypeText("[WARNING] WMI Registry data not found!")
        $selection.TypeParagraph()
        }
        elseif(($xml.info.win32_registry.errorcode) -eq 1){
        Write-Output "[WARNING] WMI Registry details not collected. Moving on to next section..."
        $warnings ++
        $selection.Style = 'Normal'
        $selection.Font.Color="255"
        $selection.TypeText("[WARNING] WMI Registry data not collected!")
        $selection.TypeParagraph()
        $selection.TypeText("Reason for error: $($xml.info.win32_registry.errortext)")
        $selection.TypeParagraph()
        }
        else{
        Write-Output "[INFO] Populating WMI Registry table."

        $table = $selection.Tables.add(
        $selection.Range,
        3,
        2,
        [Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior,
        [Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitWindow
        )

        $table.style = "Grid Table 4 - Accent 1"
        $table.cell(1,1).range.text = "Item"
        $table.cell(1,2).range.text = "Value"

        $table.cell(2,1).range.text = "Current Size (MB)"
        $table.cell(2,2).range.text = $xml.info.win32_registry.currentsize
        $table.cell(3,1).range.text = "Maximum Size (MB)"
        $table.cell(3,2).range.text = $xml.info.win32_registry.maximumsize

        $table.Rows.item(1).Headingformat=-1
        $table.ApplyStyleFirstColumn = $false
        $selection.InsertCaption(-2, ": [WMI] Registry", $null, 1, $false)

        }

    $selection.EndOf(15) | Out-Null
    $selection.MoveDown() | Out-Null
    $selection.InsertNewPage()
    #endregion ReportWMIRegistry

    #region ReportWMIEnvironment
    # BASIC
    $selection.style = 'Heading 1'
    $selection.TypeText("[WMI] Environment")
    $selection.TypeParagraph()
    $selection.TypeParagraph()

    $selection.Style = 'Normal'
    $selection.TypeText("This data is collected from the Win32_Environment WMI class. For more information please go to https://docs.microsoft.com/en-gb/windows/desktop/CIMWin32Prov/win32-environment")
    $selection.TypeParagraph()
    $selection.TypeParagraph()

        if(($xml.info.win32_environment).count -eq 0){
        Write-Output "[WARNING] WMI Environment details not found. Moving on to next section..."
        $warnings ++
        $selection.Style = 'Normal'
        $selection.Font.Color="255"
        $selection.TypeText("[WARNING] WMI Environment data not found!")
        $selection.TypeParagraph()
        }
        elseif(($xml.info.win32_environment.errorcode) -eq 1){
        Write-Output "[WARNING] WMI Environment details not collected. Moving on to next section..."
        $warnings ++
        $selection.Style = 'Normal'
        $selection.Font.Color="255"
        $selection.TypeText("[WARNING] WMI Environment data not collected!")
        $selection.TypeParagraph()
        $selection.TypeText("Reason for error: $($xml.info.win32_environment.errortext)")
        $selection.TypeParagraph()
        }
        else{
        Write-Output "[INFO] Populating WMI Environment table."

        # Count rows in multi
        $multis = $xml.selectnodes("//info/win32_environment/win32_environment_multi")
        $count = ($multis | Measure-Object).Count

            # Calculate rows. (No of items x no of multis) + (no of multis + 1 (for header)) | or less 1 for header if single row
            if($count -eq 1){
            $rows = (5 + $count)
            }
            else{
            $rows = (5 * $count) + ($count + 1)
            }

        $table = $selection.Tables.add(
        $selection.Range,
        $rows,
        2,
        [Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior,
        [Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitWindow
        )

        $table.style = "Grid Table 4 - Accent 1"
        $table.cell(1,1).range.text = "Item"
        $table.cell(1,2).range.text = "Value"

        $i = 2 # 2 as header is row 1
        $y = 1

            foreach($multi in $multis){

                if($count -gt 1){
                $table.cell($i,1).Merge($table.cell($i, 2))
                $table.cell($i,1).range.text = "------ BLOCK $y ------"
                $i++
                }

            $table.cell($i,1).range.text = "Name"
            $table.cell($i,2).range.text = $multi.name
            $i++

            $table.cell($i,1).range.text = "System Variable"
            $table.cell($i,2).range.text = $multi.systemvariable
            $i++

            $table.cell($i,1).range.text = "Caption"
            $table.cell($i,2).range.text = $multi.caption
            $i++

            $table.cell($i,1).range.text = "User Name"
            $table.cell($i,2).range.text = $multi.username
            $i++

            $table.cell($i,1).range.text = "Variable Value"
            $table.cell($i,2).range.text = $multi.variablevalue -replace ";", "`r`n"
            $i++

            $y++

            }

        $table.Rows.item(1).Headingformat=-1
        $table.ApplyStyleFirstColumn = $false
        $selection.InsertCaption(-2, ": [WMI] Environment", $null, 1, $false)

        }

    $selection.EndOf(15) | Out-Null
    $selection.MoveDown() | Out-Null
    $selection.InsertNewPage()
    #endregion ReportWMIEnvironment

    #region ReportWMICDRomDrive
    # BASIC
    $selection.style = 'Heading 1'
    $selection.TypeText("[WMI] CD ROM Drive")
    $selection.TypeParagraph()
    $selection.TypeParagraph()

    $selection.Style = 'Normal'
    $selection.TypeText("This data is collected from the Win32_CDROMDrive WMI class. For more information please go to https://docs.microsoft.com/en-gb/windows/desktop/CIMWin32Prov/win32-cdromdrive")
    $selection.TypeParagraph()
    $selection.TypeParagraph()

        if(($xml.info.win32_cdromdrive).count -eq 0){
        Write-Output "[WARNING] WMI CD ROM Drive details not found. Moving on to next section..."
        $warnings ++
        $selection.Style = 'Normal'
        $selection.Font.Color="255"
        $selection.TypeText("[WARNING] WMI CD ROM Drive data not found!")
        $selection.TypeParagraph()
        }
        elseif(($xml.info.win32_cdromdrive.errorcode) -eq 1){
        Write-Output "[WARNING] WMI CD ROM Drive details not collected. Moving on to next section..."
        $warnings ++
        $selection.Style = 'Normal'
        $selection.Font.Color="255"
        $selection.TypeText("[WARNING] WMI CD ROM Drive data not collected!")
        $selection.TypeParagraph()
        $selection.TypeText("Reason for error: $($xml.info.win32_cdromdrive.errortext)")
        $selection.TypeParagraph()
        }
        else{
        Write-Output "[INFO] Populating WMI CD ROM Drive table."

        # Count rows in multi
        $multis = $xml.selectnodes("//info/win32_cdromdrive/win32_cdromdrive_multi")
        $count = ($multis | Measure-Object).Count

            # Calculate rows. (No of items x no of multis) + (no of multis + 1 (for header)) | or less 1 for header if single row
            if($count -eq 0){
            $rows = $count + 2
            }
            elseif($count -eq 1){
            $rows = (10 + $count)
            }
            else{
            $rows = (10 * $count) + ($count + 1)
            }

        $table = $selection.Tables.add(
        $selection.Range,
        $rows,
        2,
        [Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior,
        [Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitWindow
        )

        $table.style = "Grid Table 4 - Accent 1"
        $table.cell(1,1).range.text = "Item"
        $table.cell(1,2).range.text = "Value"

        $i = 2 # 2 as header is row 1
        $y = 1

            if($count -eq 0){
            $table.cell($i,1).range.text = "No CD ROM drive installed"
            $table.Cell($i,2).Merge($table.Cell($i,1))
            }

            foreach($multi in $multis){

                if($count -gt 1){
                $table.cell($i,1).Merge($table.cell($i, 2))
                $table.cell($i,1).range.text = "------ BLOCK $y ------"
                $i++
                }

            $table.cell($i,1).range.text = "Availability"

        	    if($multi.availability -eq 1){
	            $availability = "Other"
	            }
	            elseif($multi.availability -eq 2){
	            $availability = "Unknown"
	            }
	            elseif($multi.availability -eq 3){
	            $availability = "Running/Full Power"
	            }
	            elseif($multi.availability -eq 4){
	            $availability = "Warning"
	            }
	            elseif($multi.availability -eq 5){
	            $availability = "In Test"
	            }
	            elseif($multi.availability -eq 6){
	            $availability = "Not Applicable"
	            }
	            elseif($multi.availability -eq 7){
	            $availability = "Power Off"
	            }
	            elseif($multi.availability -eq 8){
	            $availability = "Off Line"
	            }
	            elseif($multi.availability -eq 9){
	            $availability = "Off Duty"
	            }
	            elseif($multi.availability -eq 10){
	            $availability = "Degraded"
	            }
	            elseif($multi.availability -eq 11){
	            $availability = "Not Installed"
	            }
	            elseif($multi.availability -eq 12){
	            $availability = "Install Error"
	            }
	            elseif($multi.availability -eq 13){
	            $availability = "Power Save - Unknown"
	            }
	            elseif($multi.availability -eq 14){
	            $availability = "Power Save - Low Power Mode"
	            }
	            elseif($multi.availability -eq 15){
	            $availability = "Power Save - Standby"
	            }
	            elseif($multi.availability -eq 16){
	            $availability = "Power Cycle"
	            }
	            elseif($multi.availability -eq 17){
	            $availability = "Power Save - Warning"
	            }
	            elseif($multi.availability -eq 18){
	            $availability = "Paused"
	            }
	            elseif($multi.availability -eq 19){
	            $availability = "Not Ready"
	            }
	            elseif($multi.availability -eq 20){
	            $availability = "Not Configured"
	            }
	            elseif($multi.availability -eq 21){
	            $availability = "Quiesced"
	            }
	            else{
	            $availability = "$multi.availability [UNKNOWN]"
	            }

            $table.cell($i,2).range.text = $availability
            $i++

            $table.cell($i,1).range.text = "Capability Descriptions"
            $table.cell($i,2).range.text = ($multi.capabilitydescriptions -replace "  ","@") -replace "@", "`r`n" # Spacing between entries is double spaced.
            $i++

            $table.cell($i,1).range.text = "Description"
            $table.cell($i,2).range.text = $multi.description
            $i++

            $table.cell($i,1).range.text = "Device ID"
            $table.cell($i,2).range.text = $multi.deviceid
            $i++

            $table.cell($i,1).range.text = "Drive"
            $table.cell($i,2).range.text = $multi.drive
            $i++

            $table.cell($i,1).range.text = "Manufacturer"
            $table.cell($i,2).range.text = $multi.manufacturer
            $i++

            $table.cell($i,1).range.text = "Media Type"
            $table.cell($i,2).range.text = $multi.mediatype
            $i++

            $table.cell($i,1).range.text = "Name"
            $table.cell($i,2).range.text = $multi.name
            $i++

            $table.cell($i,1).range.text = "PNP Device ID"
            $table.cell($i,2).range.text = $multi.pnpdeviceid
            $i++

            $table.cell($i,1).range.text = "Serial Number"
            $table.cell($i,2).range.text = $multi.serialnumber
            $i++
            $y++

            }

        $table.Rows.item(1).Headingformat=-1
        $table.ApplyStyleFirstColumn = $false
        $selection.InsertCaption(-2, ": [WMI] CD ROM Drive", $null, 1, $false)

        }

    $selection.EndOf(15) | Out-Null
    $selection.MoveDown() | Out-Null
    $selection.InsertNewPage()
    #endregion ReportWMICDRomDrive

    #region ReportWMIVideoController
    # BASIC
    $selection.style = 'Heading 1'
    $selection.TypeText("[WMI] Video Controller")
    $selection.TypeParagraph()
    $selection.TypeParagraph()

    $selection.Style = 'Normal'
    $selection.TypeText("This data is collected from the Win32_VideoController WMI class. For more information please go to https://docs.microsoft.com/en-gb/windows/desktop/CIMWin32Prov/win32-videocontroller")
    $selection.TypeParagraph()
    $selection.TypeParagraph()

        if(($xml.info.win32_videocontroller).count -eq 0){
        Write-Output "[WARNING] WMI Video Controller details not found. Moving on to next section..."
        $warnings ++
        $selection.Style = 'Normal'
        $selection.Font.Color="255"
        $selection.TypeText("[WARNING] WMI Video Controller data not found!")
        $selection.TypeParagraph()
        }
        elseif(($xml.info.win32_videocontroller.errorcode) -eq 1){
        Write-Output "[WARNING] WMI Video Controller details not collected. Moving on to next section..."
        $warnings ++
        $selection.Style = 'Normal'
        $selection.Font.Color="255"
        $selection.TypeText("[WARNING] WMI Video Controller data not collected!")
        $selection.TypeParagraph()
        $selection.TypeText("Reason for error: $($xml.info.win32_videocontroller.errortext)")
        $selection.TypeParagraph()
        }
        else{
        Write-Output "[INFO] Populating WMI Video Controller table."

        # Count rows in multi
        $multis = $xml.selectnodes("//info/win32_videocontroller/win32_videocontroller_multi")
        $count = ($multis | Measure-Object).Count

            # Calculate rows. (No of items x no of multis) + (no of multis + 1 (for header)) | or less 1 for header if single row
            if($count -eq 0){
            $rows = $count + 2
            }
            elseif($count -eq 1){
            $rows = (14 + $count)
            }
            else{
            $rows = (14 * $count) + ($count + 1)
            }

        $table = $selection.Tables.add(
        $selection.Range,
        $rows,
        2,
        [Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior,
        [Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitWindow
        )

        $table.style = "Grid Table 4 - Accent 1"
        $table.cell(1,1).range.text = "Item"
        $table.cell(1,2).range.text = "Value"

        $i = 2 # 2 as header is row 1
        $y = 1

            if($count -eq 0){
            $table.cell($i,1).range.text = "No video controllers installed"
            $table.Cell($i,2).Merge($table.Cell($i,1))
            }

            foreach($multi in $multis){

                if($count -gt 1){
                $table.cell($i,1).Merge($table.cell($i, 2))
                $table.cell($i,1).range.text = "------ BLOCK $y ------"
                $i++
                }

            $table.cell($i,1).range.text = "Adapter Compatibility"
            $table.cell($i,2).range.text = $multi.adaptercompatibility
            $i++

            $table.cell($i,1).range.text = "Adapter DAC Type"
            $table.cell($i,2).range.text = $multi.adapterdactype
            $i++

            # Convert bytes to GB and Round
            [string]$a = [system.decimal](($multi.adapterram)/1073741824)
            $b = [math]::Round($a,2)

            $table.cell($i,1).range.text = "Adapter RAM (GB)"
            $table.cell($i,2).range.text = "$b"
            $i++

            $table.cell($i,1).range.text = "Availability"

                if($multi.availability -eq 1){
	            $availability = "Other"
	            }
	            elseif($multi.availability -eq 2){
	            $availability = "Unknown"
	            }
	            elseif($multi.availability -eq 3){
	            $availability = "Running/Full Power"
	            }
	            elseif($multi.availability -eq 4){
	            $availability = "Warning"
	            }
	            elseif($multi.availability -eq 5){
	            $availability = "In Test"
	            }
	            elseif($multi.availability -eq 6){
	            $availability = "Not Applicable"
	            }
	            elseif($multi.availability -eq 7){
	            $availability = "Power Off"
	            }
	            elseif($multi.availability -eq 8){
	            $availability = "Off Line"
	            }
	            elseif($multi.availability -eq 9){
	            $availability = "Off Duty"
	            }
	            elseif($multi.availability -eq 10){
	            $availability = "Degraded"
	            }
	            elseif($multi.availability -eq 11){
	            $availability = "Not Installed"
	            }
	            elseif($multi.availability -eq 12){
	            $availability = "Install Error"
	            }
	            elseif($multi.availability -eq 13){
	            $availability = "Power Save - Unknown"
	            }
	            elseif($multi.availability -eq 14){
	            $availability = "Power Save - Low Power Mode"
	            }
	            elseif($multi.availability -eq 15){
	            $availability = "Power Save - Standby"
	            }
	            elseif($multi.availability -eq 16){
	            $availability = "Power Cycle"
	            }
	            elseif($multi.availability -eq 17){
	            $availability = "Power Save - Warning"
	            }
	            elseif($multi.availability -eq 18){
	            $availability = "Paused"
	            }
	            elseif($multi.availability -eq 19){
	            $availability = "Not Ready"
	            }
	            elseif($multi.availability -eq 20){
	            $availability = "Not Configured"
	            }
	            elseif($multi.availability -eq 21){
	            $availability = "Quiesced"
	            }
	            else{
	            $availability = "$multi.availability [UNKNOWN]"
	            }

            $table.cell($i,2).range.text = $availability
            $i++

            $table.cell($i,1).range.text = "Device ID"
            $table.cell($i,2).range.text = $multi.deviceid
            $i++

            $driverDate =
            $multi.driverdate.Substring(6,2) + "/" +
            $multi.driverdate.Substring(4,2) + "/" +
            $multi.driverdate.Substring(0,4)

            $table.cell($i,1).range.text = "Driver Date"
            $table.cell($i,2).range.text = $driverDate
            $i++

            $table.cell($i,1).range.text = "Driver Version"
            $table.cell($i,2).range.text = $multi.driverversion
            $i++

            $table.cell($i,1).range.text = "INF File Name"
            $table.cell($i,2).range.text = $multi.inffilename
            $i++

            $table.cell($i,1).range.text = "INF Section"
            $table.cell($i,2).range.text = $multi.infsection
            $i++

            $table.cell($i,1).range.text = "Installed Display Drivers"
            $table.cell($i,2).range.text = $multi.installeddisplaydrivers -replace ",", "`r`n"
            $i++

            $table.cell($i,1).range.text = "Name"
            $table.cell($i,2).range.text = $multi.name
            $i++

            $table.cell($i,1).range.text = "PNP Device ID"
            $table.cell($i,2).range.text = $multi.pnpdeviceid
            $i++

            $table.cell($i,1).range.text = "Video Mode Description"
            $table.cell($i,2).range.text = $multi.videomodedescription
            $i++

            $table.cell($i,1).range.text = "Video Processor"
            $table.cell($i,2).range.text = $multi.videoprocessor
            $i++

            $y++

            }

        $table.Rows.item(1).Headingformat=-1
        $table.ApplyStyleFirstColumn = $false
        $selection.InsertCaption(-2, ": [WMI] Video Controller", $null, 1, $false)

        }

    $selection.EndOf(15) | Out-Null
    $selection.MoveDown() | Out-Null
    $selection.InsertNewPage()
    #endregion ReportWMIVideoController

    #region ReportWMISoundDevice
    # BASIC
    $selection.style = 'Heading 1'
    $selection.TypeText("[WMI] Sound Device")
    $selection.TypeParagraph()
    $selection.TypeParagraph()

    $selection.Style = 'Normal'
    $selection.TypeText("This data is collected from the Win32_SoundDevice WMI class. For more information please go to https://docs.microsoft.com/en-gb/windows/desktop/CIMWin32Prov/win32-sounddevice")
    $selection.TypeParagraph()
    $selection.TypeParagraph()

        if(($xml.info.win32_sounddevice).count -eq 0){
        Write-Output "[WARNING] WMI Sound Device details not found. Moving on to next section..."
        $warnings ++
        $selection.Style = 'Normal'
        $selection.Font.Color="255"
        $selection.TypeText("[WARNING] WMI Sound Device data not found!")
        $selection.TypeParagraph()
        }
        elseif(($xml.info.win32_sounddevice.errorcode) -eq 1){
        Write-Output "[WARNING] WMI Sound Device details not collected. Moving on to next section..."
        $warnings ++
        $selection.Style = 'Normal'
        $selection.Font.Color="255"
        $selection.TypeText("[WARNING] WMI Sound Device data not collected!")
        $selection.TypeParagraph()
        $selection.TypeText("Reason for error: $($xml.info.win32_sounddevice.errortext)")
        $selection.TypeParagraph()
        }
        else{
        Write-Output "[INFO] Populating WMI Sound Device table."

        # Count rows in multi
        $multis = $xml.selectnodes("//info/win32_sounddevice/win32_sounddevice_multi")
        $count = ($multis | Measure-Object).Count

            # Calculate rows. (No of items x no of multis) + (no of multis + 1 (for header)) | or less 1 for header if single row
            if($count -eq 0){
            $rows = $count + 2
            }
            elseif($count -eq 1){
            $rows = (6 + $count)
            }
            else{
            $rows = (6 * $count) + ($count + 1)
            }

        $table = $selection.Tables.add(
        $selection.Range,
        $rows,
        2,
        [Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior,
        [Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitWindow
        )

        $table.style = "Grid Table 4 - Accent 1"
        $table.cell(1,1).range.text = "Item"
        $table.cell(1,2).range.text = "Value"

        $i = 2 # 2 as header is row 1
        $y = 1

            if($count -eq 0){
            $table.cell($i,1).range.text = "No sound devices installed"
            $table.Cell($i,2).Merge($table.Cell($i,1))
            }

            foreach($multi in $multis){

                if($count -gt 1){
                $table.cell($i,1).Merge($table.cell($i, 2))
                $table.cell($i,1).range.text = "------ BLOCK $y ------"
                $i++
                }

            $table.cell($i,1).range.text = "Device ID"
            $table.cell($i,2).range.text = $multi.deviceid
            $i++

            $table.cell($i,1).range.text = "Manufacturer"
            $table.cell($i,2).range.text = $multi.manufacturer
            $i++

            $table.cell($i,1).range.text = "Name"
            $table.cell($i,2).range.text = $multi.name
            $i++

            $table.cell($i,1).range.text = "PNP Device ID"
            $table.cell($i,2).range.text = $multi.pnpdeviceid
            $i++

            $table.cell($i,1).range.text = "Product Name"
            $table.cell($i,2).range.text = $multi.productname
            $i++

            $table.cell($i,1).range.text = "Status Info"

                if($multi.statusinfo -eq 1){
                $statusInfo = "Other"
                }
                elseif($multi.statusinfo -eq 2){
                $statusInfo = "Unknown"
                }
                elseif($multi.statusinfo -eq 3){
                $statusInfo = "Enabled"
                }
                elseif($multi.statusinfo -eq 4){
                $statusInfo = "Disabled"
                }
                elseif($multi.statusinfo -eq 5){
                $statusInfo = "Not Available"
                }
                else{
                $statusInfo = "$multi.statusinfo [UKNOWN]"
                }

            $table.cell($i,2).range.text = $statusinfo
            $i++

            $y++

            }

        $table.Rows.item(1).Headingformat=-1
        $table.ApplyStyleFirstColumn = $false
        $selection.InsertCaption(-2, ": [WMI] Sound Device", $null, 1, $false)

        }

    $selection.EndOf(15) | Out-Null
    $selection.MoveDown() | Out-Null
    $selection.InsertNewPage()
    #endregion ReportWMISoundDevice

    #region ReportWMIPrinters
    # BASIC & DETAILED
    $selection.style = 'Heading 1'
    $selection.TypeText("[WMI] Printers")
    $selection.TypeParagraph()
    $selection.TypeParagraph()

    $selection.Style = 'Normal'
    $selection.TypeText("This data is collected from the Win32_Printer WMI class. For more information please go to https://docs.microsoft.com/en-gb/windows/desktop/CIMWin32Prov/win32-printer")
    $selection.TypeParagraph()
    $selection.TypeParagraph()


        if($reportType -ne 'Basic'){
        $selection.style = 'Heading 2'
        $selection.TypeText("Basic")
        $selection.TypeParagraph()
        $selection.TypeParagraph()
        }

        if(($xml.info.win32_printer).count -eq 0){
        Write-Output "[WARNING] WMI Printer (Basic) details not found. Moving on to next section..."
        $warnings ++
        $selection.Style = 'Normal'
        $selection.Font.Color="255"
        $selection.TypeText("[WARNING] WMI Printer (Basic) data not found!")
        $selection.TypeParagraph()
        }
        elseif(($xml.info.win32_printer.errorcode) -eq 1){
        Write-Output "[WARNING] WMI Printer (Basic) details not collected. Moving on to next section..."
        $warnings ++
        $selection.Style = 'Normal'
        $selection.Font.Color="255"
        $selection.TypeText("[WARNING] WMI Printer (Basic) data not collected!")
        $selection.TypeParagraph()
        $selection.TypeText("Reason for error: $($xml.info.win32_printer.errortext)")
        $selection.TypeParagraph()
        }
        else{
        Write-Output "[INFO] Populating WMI Printer (Basic) table."

        # Count rows in multi
        $multis = $xml.selectnodes("//info/win32_printer/win32_printer_multi")
        $count = ($multis | Measure-Object).Count

            # Calculate rows (from query above) and add 1 for the header
            if($count -eq 0){
            $rows = $count + 2
            }
            else{
            $rows = $count + 1
            }

        $table = $selection.Tables.add(
        $selection.Range,
        $rows,
        4,
        [Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior,
        [Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitWindow
        )

        $table.style = "Grid Table 4 - Accent 1"
        $table.cell(1,1).range.text = "Name"
        $table.cell(1,2).range.text = "Local"
        $table.cell(1,3).range.text = "Network"
        $table.cell(1,4).range.text = "Shared"

        $i = 2 # 2 as header is row 1

            if($count -eq 0){
            $table.cell($i,1).range.text = "No printers installed"
            $table.Cell($i,2).Merge($table.Cell($i,1))
            $table.Cell($i,2).Merge($table.Cell($i,1))
            $table.Cell($i,2).Merge($table.Cell($i,1))
            }

            foreach($multi in $multis){

            $table.cell($i,1).range.text = $multi.name
            $table.cell($i,2).range.text = $multi.local
            $table.cell($i,3).range.text = $multi.network
            $table.cell($i,4).range.text = $multi.shared
            $i++

            }

        $table.Rows.item(1).Headingformat=-1
        $table.ApplyStyleFirstColumn = $false
        $selection.InsertCaption(-2, ": [WMI] Printer (Basic)", $null, 1, $false)

          if($reportType -ne "Basic"){
          $selection.TypeParagraph()
          }

        }

    $selection.EndOf(15) | Out-Null
    $selection.MoveDown() | Out-Null

        if($reportType -eq "Detailed"){
        $selection.TypeParagraph()
        $selection.style = 'Heading 2'
        $selection.TypeText("Detailed")
        $selection.TypeParagraph()
        $selection.TypeParagraph()

            if(($xml.info.win32_printer).count -eq 0){
            Write-Output "[WARNING] WMI Printer (Detailed) details not found. Moving on to next section..."
            $warnings ++
            $selection.Style = 'Normal'
            $selection.Font.Color="255"
            $selection.TypeText("[WARNING] WMI Printer (Detailed) data not found!")
            $selection.TypeParagraph()
            }
            elseif(($xml.info.win32_printer.errorcode) -eq 1){
            Write-Output "[WARNING] WMI Printer (Detailed) details not collected. Moving on to next section..."
            $warnings ++
            $selection.Style = 'Normal'
            $selection.Font.Color="255"
            $selection.TypeText("[WARNING] WMI Printer (Detailed) data not collected!")
            $selection.TypeParagraph()
            $selection.TypeText("Reason for error: $($xml.info.win32_printer.errortext)")
            $selection.TypeParagraph()
            }
            else{
            Write-Output "[INFO] Populating WMI Printer (Detailed) table."

            # Count rows in multi
            $multis = $xml.selectnodes("//info/win32_printer/win32_printer_multi")
            $count = ($multis | Measure-Object).Count

                # Calculate rows. (No of items x no of multis) + (no of multis + 1 (for header)) | or less 1 for header if single row
                if($count -eq 0){
                $rows = $count + 2
                }
                elseif($count -eq 1){
                $rows = (14 + $count)
                }
                else{
                $rows = (14 * $count) + ($count + 1)
                }

            $table = $selection.Tables.add(
            $selection.Range,
            $rows,
            2,
            [Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior,
            [Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitWindow
            )

            $table.style = "Grid Table 4 - Accent 1"
            $table.cell(1,1).range.text = "Item"
            $table.cell(1,2).range.text = "Value"

            $i = 2 # 2 as header is row 1
            $y = 1

                if($count -eq 0){
                $table.cell($i,1).range.text = "No printers installed"
                $table.Cell($i,2).Merge($table.Cell($i,1))
                }

                foreach($multi in $multis){

                    if($count -gt 1){
                    $table.cell($i,1).Merge($table.cell($i, 2))
                    $table.cell($i,1).range.text = "------ BLOCK $y ------"
                    $i++
                    }

                $table.cell($i,1).range.text = "Capability Descriptions"
                $table.cell($i,2).range.text = $multi.capabilitydescriptions
                $i++

                $table.cell($i,1).range.text = "Default"
                $table.cell($i,2).range.text = $multi.default
                $i++

                $table.cell($i,1).range.text = "Device ID"
                $table.cell($i,2).range.text = $multi.deviceid
                $i++

                $table.cell($i,1).range.text = "Driver Name"
                $table.cell($i,2).range.text = $multi.drivername
                $i++

                $table.cell($i,1).range.text = "Hidden"
                $table.cell($i,2).range.text = $multi.hidden
                $i++

                $table.cell($i,1).range.text = "Local"
                $table.cell($i,2).range.text = $multi.local
                $i++

                $table.cell($i,1).range.text = "Name"
                $table.cell($i,2).range.text = $multi.name
                $i++

                $table.cell($i,1).range.text = "Network"
                $table.cell($i,2).range.text = $multi.network
                $i++

                $table.cell($i,1).range.text = "PNP Device ID"
                $table.cell($i,2).range.text = $multi.pnpdeviceid
                $i++

                $table.cell($i,1).range.text = "Port Name"
                $table.cell($i,2).range.text = $multi.portname
                $i++

                $table.cell($i,1).range.text = "Print Processor"
                $table.cell($i,2).range.text = $multi.printprocessor
                $i++

                $table.cell($i,1).range.text = "Server Name"
                $table.cell($i,2).range.text = $multi.servername
                $i++

                $table.cell($i,1).range.text = "Share Name"
                $table.cell($i,2).range.text = $multi.sharename
                $i++

                $table.cell($i,1).range.text = "Shared"
                $table.cell($i,2).range.text = $multi.shared
                $i++

                $y++

                }

            $table.Rows.item(1).Headingformat=-1
            $table.ApplyStyleFirstColumn = $false
            $selection.InsertCaption(-2, ": [WMI] Printer (Detailed)", $null, 1, $false)

            }

        }

    $selection.EndOf(15) | Out-Null
    $selection.MoveDown() | Out-Null
    $selection.InsertNewPage()
    #endregion ReportWMIPrinters

    #region ReportWMIDiskDrive
    # BASIC
    $selection.style = 'Heading 1'
    $selection.TypeText("[WMI] Disk Drive")
    $selection.TypeParagraph()
    $selection.TypeParagraph()

    $selection.Style = 'Normal'
    $selection.TypeText("This data is collected from the Win32_DiskDrive WMI class. For more information please go to https://docs.microsoft.com/en-gb/windows/desktop/CIMWin32Prov/win32-diskdrive")
    $selection.TypeParagraph()
    $selection.TypeParagraph()

        if(($xml.info.win32_diskdrive).count -eq 0){
        Write-Output "[WARNING] WMI Disk Drive details not found. Moving on to next section..."
        $warnings ++
        $selection.Style = 'Normal'
        $selection.Font.Color="255"
        $selection.TypeText("[WARNING] WMI Disk Drive data not found!")
        $selection.TypeParagraph()
        }
        elseif(($xml.info.win32_diskdrive.errorcode) -eq 1){
        Write-Output "[WARNING] WMI Disk Drive details not collected. Moving on to next section..."
        $warnings ++
        $selection.Style = 'Normal'
        $selection.Font.Color="255"
        $selection.TypeText("[WARNING] WMI Disk Drive data not collected!")
        $selection.TypeParagraph()
        $selection.TypeText("Reason for error: $($xml.info.win32_diskdrive.errortext)")
        $selection.TypeParagraph()
        }
        else{
        Write-Output "[INFO] Populating WMI Disk Drive table."

        # Count rows in multi
        $multis = $xml.selectnodes("//info/win32_diskdrive/win32_diskdrive_multi")
        $count = ($multis | Measure-Object).Count

            # Calculate rows. (No of items x no of multis) + (no of multis + 1 (for header)) | or less 1 for header if single row
            if($count -eq 1){
            $rows = (10 + $count)
            }
            else{
            $rows = (10 * $count) + ($count + 1)
            }

        $table = $selection.Tables.add(
        $selection.Range,
        $rows,
        2,
        [Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior,
        [Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitWindow
        )

        $table.style = "Grid Table 4 - Accent 1"
        $table.cell(1,1).range.text = "Item"
        $table.cell(1,2).range.text = "Value"

        $i = 2 # 2 as header is row 1
        $y = 1

            foreach($multi in $multis){

                if($count -gt 1){
                $table.cell($i,1).Merge($table.cell($i, 2))
                $table.cell($i,1).range.text = "------ BLOCK $y ------"
                $i++
                }

            $table.cell($i,1).range.text = "Caption"
            $table.cell($i,2).range.text = $multi.caption
            $i++

            $table.cell($i,1).range.text = "Firmware Revision"
            $table.cell($i,2).range.text = $multi.firmwarerevision
            $i++

            $table.cell($i,1).range.text = "Interface Type"
            $table.cell($i,2).range.text = $multi.interfacetype
            $i++

            $table.cell($i,1).range.text = "Manufacturer"
            $table.cell($i,2).range.text = $multi.manufacturer
            $i++

            $table.cell($i,1).range.text = "Model"
            $table.cell($i,2).range.text = $multi.model
            $i++

            $table.cell($i,1).range.text = "Name"
            $table.cell($i,2).range.text = $multi.name
            $i++

            $table.cell($i,1).range.text = "PNP Device ID"
            $table.cell($i,2).range.text = $multi.pnpdeviceid
            $i++

            $table.cell($i,1).range.text = "Partitions"
            $table.cell($i,2).range.text = $multi.partitions
            $i++

            $table.cell($i,1).range.text = "Serial Number"
            $table.cell($i,2).range.text = $multi.serialnumber
            $i++

            $table.cell($i,1).range.text = "Size (GB)"

            # Convert bytes to GB and Round
            [string]$a = [system.decimal](($multi.size)/1073741824)
            $b = [math]::Round($a,2)

            $table.cell($i,2).range.text = "$b"
            $i++

            $y++

            }

        $table.Rows.item(1).Headingformat=-1
        $table.ApplyStyleFirstColumn = $false
        $selection.InsertCaption(-2, ": [WMI] Disk Drive", $null, 1, $false)

        }

    $selection.EndOf(15) | Out-Null
    $selection.MoveDown() | Out-Null
    $selection.InsertNewPage()
    #endregion ReportWMIDiskDrive

    #region ReportWMIEncryptableVolume
    # BASIC
    $selection.style = 'Heading 1'
    $selection.TypeText("[WMI] Encryptable Volume")
    $selection.TypeParagraph()
    $selection.TypeParagraph()

    $selection.Style = 'Normal'
    $selection.TypeText("This data is collected from the Win32_EncryptableVolume WMI class. For more information please go to https://docs.microsoft.com/en-us/windows/win32/secprov/win32-encryptablevolume")
    $selection.TypeParagraph()
    $selection.TypeParagraph()

        if(($xml.info.win32_encryptablevolume).count -eq 0){
        Write-Output "[WARNING] WMI Encryptable Volume details not found. Moving on to next section..."
        $warnings ++
        $selection.Style = 'Normal'
        $selection.Font.Color="255"
        $selection.TypeText("[WARNING] WMI Encryptable Volume data not found!")
        $selection.TypeParagraph()
        }
        elseif(($xml.info.win32_encryptablevolume.errorcode) -eq 1){
        Write-Output "[WARNING] WMI Encryptable Volume details not collected. Moving on to next section..."
        $warnings ++
        $selection.Style = 'Normal'
        $selection.Font.Color="255"
        $selection.TypeText("[WARNING] WMI Encryptable Volume data not collected!")
        $selection.TypeParagraph()
        $selection.TypeText("Reason for error: $($xml.info.win32_encryptablevolume.errortext)")
        $selection.TypeParagraph()
        }
        else{
        Write-Output "[INFO] Populating WMI Encryptable Volume table."

        # Count rows in multi
        $multis = $xml.selectnodes("//info/win32_encryptablevolume/win32_encryptablevolume_multi")
        $count = ($multis | Measure-Object).Count

            # Calculate rows. (No of items x no of multis) + (no of multis + 1 (for header)) | or less 1 for header if single row
            if($count -eq 0){
            $rows = $count + 2
            }
            elseif($count -eq 1){
            $rows = (6 + $count)
            }
            else{
            $rows = (6 * $count) + ($count + 1)
            }

        $table = $selection.Tables.add(
        $selection.Range,
        $rows,
        2,
        [Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior,
        [Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitWindow
        )

        $table.style = "Grid Table 4 - Accent 1"
        $table.cell(1,1).range.text = "Item"
        $table.cell(1,2).range.text = "Value"

        $i = 2 # 2 as header is row 1
        $y = 1

            if($count -eq 0){
            $table.cell($i,1).range.text = "No encryptable volumes found"
            $table.Cell($i,2).Merge($table.Cell($i,1))
            }

            foreach($multi in $multis){

                if($count -gt 1){
                $table.cell($i,1).Merge($table.cell($i, 2))
                $table.cell($i,1).range.text = "------ BLOCK $y ------"
                $i++
                }

            $table.cell($i,1).range.text = "Drive Letter"
            $table.cell($i,2).range.text = $multi.driveletter
            $i++

            $table.cell($i,1).range.text = "Conversion Status"

                if($multi.conversionstatus -eq 0){
                $conversionStatus = "Fully Decrypted"
                }
                elseif($multi.conversionstatus -eq 1){
                $conversionStatus = "Fully Encrypted"
                }
                elseif($multi.conversionstatus -eq 2){
                $conversionStatus = "Encryption In Progress"
                }
                elseif($multi.conversionstatus -eq 3){
                $conversionStatus = "Decryption In Progress"
                }
                elseif($multi.conversionstatus -eq 4){
                $conversionStatus = "Encryption Paused"
                }
                elseif($multi.conversionstatus -eq 5){
                $conversionStatus = "Decryption Paused"
                }
                else{
                $conversionStatus = "[UNKNOWN] Value is $multi.conversionstatus"
                }

            $table.cell($i,2).range.text = "$conversionStatus"
            $i++

            $table.cell($i,1).range.text = "Encryption Method"

                if($multi.encryptionmethod -eq 0){
                $encryptionMethod = "None"
                }
                elseif($multi.encryptionmethod -eq 1){
                $encryptionMethod = "AES_128_WITH_DIFFUSER"
                }
                elseif($multi.encryptionmethod -eq 2){
                $encryptionMethod = "AES_256_WITH_DIFFUSER"
                }
                elseif($multi.encryptionmethod -eq 3){
                $encryptionMethod = "AES_128"
                }
                elseif($multi.encryptionmethod -eq 4){
                $encryptionMethod = "AES_256"
                }
                elseif($multi.encryptionmethod -eq 5){
                $encryptionMethod = "HARDWARE_ENCRYPTION"
                }
                elseif($multi.encryptionmethod -eq 6){
                $encryptionMethod = "XTS_AES_128"
                }
                elseif($multi.encryptionmethod -eq 7){
                $encryptionMethod = "XTS_AES_256"
                }
                else{
                $encryptionMethod = "Unknown. The volume has been fully or partially encrypted with an unknown algorithm and key size"
                }

            $table.cell($i,2).range.text = "$encryptionMethod"
            $i++

            $table.cell($i,1).range.text = "Initialized For Protection"
            $table.cell($i,2).range.text = $multi.isvolumeinitializedforprotection
            $i++

            $table.cell($i,1).range.text = "Protection Status"

                if($multi.protectionstatus -eq 0){
                $protectionStatus = "Protection Off"
                }
                elseif($multi.protectionstatus -eq 1){
                $protectionStatus = "Protection On"
                }
                else{
                $protectionStatus = "Protection Unknown"
                }

            $table.cell($i,2).range.text = "$protectionstatus"
            $i++

            $table.cell($i,1).range.text = "Volume Type"

                if($multi.volumetype -eq 0){
                $volumeType = "OS Volume"
                }
                elseif($multi.volumetype -eq 1){
                $volumeType = "Fixed Data Volume"
                }
                elseif($multi.volumetype -eq 2){
                $volumeType = "Portable Date Volume"
                }
                else{
                $volumeType = "[UNKNOWN] Value is $multi.volumetype"
                }

            $table.cell($i,2).range.text = $volumeType
            $i++

            $y++

            }

        $table.Rows.item(1).Headingformat=-1
        $table.ApplyStyleFirstColumn = $false
        $selection.InsertCaption(-2, ": [WMI] Encryption Status", $null, 1, $false)

        }

    $selection.EndOf(15) | Out-Null
    $selection.MoveDown() | Out-Null
    $selection.InsertNewPage()
    #endregion ReportWMIEncryptableVolume

    #region ReportWMIFirewallProduct
    # BASIC
    $selection.style = 'Heading 1'
    $selection.TypeText("[WMI] Firewall Product")
    $selection.TypeParagraph()
    $selection.TypeParagraph()

    $selection.Style = 'Normal'
    $selection.TypeText("This data is collected from the FirewallProduct WMI class. There is no Microsoft published information on this class.")
    $selection.TypeParagraph()
    $selection.TypeParagraph()

        if(($xml.info.firewallproduct).count -eq 0){
        Write-Output "[WARNING] WMI Firewall Product details not found. Moving on to next section..."
        $warnings ++
        $selection.Style = 'Normal'
        $selection.Font.Color="255"
        $selection.TypeText("[WARNING] WMI Firewall Product data not found!")
        $selection.TypeParagraph()
        }
        elseif(($xml.info.firewallproduct.errorcode) -eq 1){
        Write-Output "[WARNING] WMI Firewall Product details not collected. Moving on to next section..."
        $warnings ++
        $selection.Style = 'Normal'
        $selection.Font.Color="255"
        $selection.TypeText("[WARNING] WMI Firewall Product data not collected!")
        $selection.TypeParagraph()
        $selection.TypeText("Reason for error: $($xml.info.firewallproduct.errortext)")
        $selection.TypeParagraph()
        }
        else{
        Write-Output "[INFO] Populating WMI Firewall Product table."

        # Count rows in multi
        $multis = $xml.selectnodes("//info/firewallproduct/firewallproduct_multi")
        $count = ($multis | Measure-Object).Count

            # Calculate rows. (No of items x no of multis) + (no of multis + 1 (for header)) | or less 1 for header if single row
            if($count -eq 0){
            $rows = $count + 2
            }
            elseif($count -eq 1){
            $rows = (4 + $count)
            }
            else{
            $rows = (4 * $count) + ($count + 1)
            }

        $table = $selection.Tables.add(
        $selection.Range,
        $rows,
        2,
        [Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior,
        [Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitWindow
        )

        $table.style = "Grid Table 4 - Accent 1"
        $table.cell(1,1).range.text = "Item"
        $table.cell(1,2).range.text = "Value"

        $i = 2 # 2 as header is row 1
        $y = 1

            if($count -eq 0){
            $table.cell($i,1).range.text = "No firewall product installed"
            $table.Cell($i,2).Merge($table.Cell($i,1))
            }

            foreach($multi in $multis){

                if($count -gt 1){
                $table.cell($i,1).Merge($table.cell($i, 2))
                $table.cell($i,1).range.text = "------ BLOCK $y ------"
                $i++
                }

            $table.cell($i,1).range.text = "Display Name"
            $table.cell($i,2).range.text = $multi.displayname
            $i++

            $table.cell($i,1).range.text = "Product EXE"
            $table.cell($i,2).range.text = $multi.pathtosignedproductexe
            $i++

            $table.cell($i,1).range.text = "Reporting EXE"
            $table.cell($i,2).range.text = $multi.pathtosignedreportingexe
            $i++

            $table.cell($i,1).range.text = "Timestamp"
            $table.cell($i,2).range.text = $multi.timestamp
            $i++

            $y++

            }

        $table.Rows.item(1).Headingformat=-1
        $table.ApplyStyleFirstColumn = $false
        $selection.InsertCaption(-2, ": [WMI] Firewall Product", $null, 1, $false)

        }

    $selection.EndOf(15) | Out-Null
    $selection.MoveDown() | Out-Null
    $selection.InsertNewPage()
    #endregion ReportWMIFirewallProduct

    #region ReportWMIAntivirusProduct
    # BASIC
    $selection.style = 'Heading 1'
    $selection.TypeText("[WMI] Antivirus Product")
    $selection.TypeParagraph()
    $selection.TypeParagraph()

    $selection.Style = 'Normal'
    $selection.TypeText("This data is collected from the AntivirusProduct WMI class. There is no Microsoft published information on this class.")
    $selection.TypeParagraph()
    $selection.TypeParagraph()

        if(($xml.info.antivirusproduct).count -eq 0){
        Write-Output "[WARNING] WMI Antivirus Product details not found. Moving on to next section..."
        $warnings ++
        $selection.Style = 'Normal'
        $selection.Font.Color="255"
        $selection.TypeText("[WARNING] WMI Antivirus Product data not found!")
        $selection.TypeParagraph()
        }
        elseif(($xml.info.antivirusproduct.errorcode) -eq 1){
        Write-Output "[WARNING] WMI Antivirus Product details not collected. Moving on to next section..."
        $warnings ++
        $selection.Style = 'Normal'
        $selection.Font.Color="255"
        $selection.TypeText("[WARNING] WMI Antivirus Product data not collected!")
        $selection.TypeParagraph()
        $selection.TypeText("Reason for error: $($xml.info.antivirusproduct.errortext)")
        $selection.TypeParagraph()
        }
        else{
        Write-Output "[INFO] Populating WMI Antivirus Product table."

        # Count rows in multi
        $multis = $xml.selectnodes("//info/antivirusproduct/antivirusproduct_multi")
        $count = ($multis | Measure-Object).Count

            # Calculate rows. (No of items x no of multis) + (no of multis + 1 (for header)) | or less 1 for header if single row
            if($count -eq 0){
            $rows = $count + 2
            }
            elseif($count -eq 1){
            $rows = (4 + $count)
            }
            else{
            $rows = (4 * $count) + ($count + 1)
            }

        $table = $selection.Tables.add(
        $selection.Range,
        $rows,
        2,
        [Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior,
        [Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitWindow
        )

        $table.style = "Grid Table 4 - Accent 1"
        $table.cell(1,1).range.text = "Item"
        $table.cell(1,2).range.text = "Value"

        $i = 2 # 2 as header is row 1
        $y = 1

            if($count -eq 0){
            $table.cell($i,1).range.text = "No antivirus product installed"
            $table.Cell($i,2).Merge($table.Cell($i,1))
            }

            foreach($multi in $multis){

                if($count -gt 1){
                $table.cell($i,1).Merge($table.cell($i, 2))
                $table.cell($i,1).range.text = "------ BLOCK $y ------"
                $i++
                }

            $table.cell($i,1).range.text = "Display Name"
            $table.cell($i,2).range.text = $multi.displayname
            $i++

            $table.cell($i,1).range.text = "Product EXE"
            $table.cell($i,2).range.text = $multi.pathtosignedproductexe
            $i++

            $table.cell($i,1).range.text = "Reporting EXE"
            $table.cell($i,2).range.text = $multi.pathtosignedreportingexe
            $i++

            $table.cell($i,1).range.text = "Timestamp"
            $table.cell($i,2).range.text = $multi.timestamp
            $i++

            $y++

            }

        $table.Rows.item(1).Headingformat=-1
        $table.ApplyStyleFirstColumn = $false
        $selection.InsertCaption(-2, ": [WMI] Antivirus Product", $null, 1, $false)

        }

    $selection.EndOf(15) | Out-Null
    $selection.MoveDown() | Out-Null
    $selection.InsertNewPage()
    #endregion ReportWMIAntivirusProduct

    #region ReportWMIAntispywareProduct
    # BASIC
    $selection.style = 'Heading 1'
    $selection.TypeText("[WMI] Antispyware Product")
    $selection.TypeParagraph()
    $selection.TypeParagraph()

    $selection.Style = 'Normal'
    $selection.TypeText("This data is collected from the AntisypwareProduct WMI class. There is no Microsoft published information on this class.")
    $selection.TypeParagraph()
    $selection.TypeParagraph()

        if(($xml.info.antispywareproduct).count -eq 0){
        Write-Output "[WARNING] WMI Antispyware Product details not found. Moving on to next section..."
        $warnings ++
        $selection.Style = 'Normal'
        $selection.Font.Color="255"
        $selection.TypeText("[WARNING] WMI Antispyware Product data not found!")
        $selection.TypeParagraph()
        }
        elseif(($xml.info.antispywareproduct.errorcode) -eq 1){
        Write-Output "[WARNING] WMI Antispyware Product details not collected. Moving on to next section..."
        $warnings ++
        $selection.Style = 'Normal'
        $selection.Font.Color="255"
        $selection.TypeText("[WARNING] WMI Antispyware Product data not collected!")
        $selection.TypeParagraph()
        $selection.TypeText("Reason for error: $($xml.info.antispywareproduct.errortext)")
        $selection.TypeParagraph()
        }
        else{
        Write-Output "[INFO] Populating WMI Antispyware Product table."

        # Count rows in multi
        $multis = $xml.selectnodes("//info/antispywareproduct/antispywareproduct_multi")
        $count = ($multis | Measure-Object).Count

            # Calculate rows. (No of items x no of multis) + (no of multis + 1 (for header)) | or less 1 for header if single row
            if($count -eq 0){
            $rows = $count + 2
            }
            elseif($count -eq 1){
            $rows = (4 + $count)
            }
            else{
            $rows = (4 * $count) + ($count + 1)
            }

        $table = $selection.Tables.add(
        $selection.Range,
        $rows,
        2,
        [Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior,
        [Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitWindow
        )

        $table.style = "Grid Table 4 - Accent 1"
        $table.cell(1,1).range.text = "Item"
        $table.cell(1,2).range.text = "Value"

        $i = 2 # 2 as header is row 1
        $y = 1

            if($count -eq 0){
            $table.cell($i,1).range.text = "No antispyware products installed"
            $table.Cell($i,2).Merge($table.Cell($i,1))
            }

            foreach($multi in $multis){

                if($count -gt 1){
                $table.cell($i,1).Merge($table.cell($i, 2))
                $table.cell($i,1).range.text = "------ BLOCK $y ------"
                $i++
                }

            $table.cell($i,1).range.text = "Display Name"
            $table.cell($i,2).range.text = $multi.displayname
            $i++

            $table.cell($i,1).range.text = "Product EXE"
            $table.cell($i,2).range.text = $multi.pathtosignedproductexe
            $i++

            $table.cell($i,1).range.text = "Reporting EXE"
            $table.cell($i,2).range.text = $multi.pathtosignedreportingexe
            $i++

            $table.cell($i,1).range.text = "Timestamp"
            $table.cell($i,2).range.text = $multi.timestamp
            $i++

            $y++

            }

        $table.Rows.item(1).Headingformat=-1
        $table.ApplyStyleFirstColumn = $false
        $selection.InsertCaption(-2, ": [WMI] Antispyware Product", $null, 1, $false)

        }

    $selection.EndOf(15) | Out-Null
    $selection.MoveDown() | Out-Null
    $selection.InsertNewPage()
    #endregion ReportWMIAntispywareProduct

    #region ReportWMIOptionalFeatures
    # BASIC
    $selection.style = 'Heading 1'
    $selection.TypeText("[WMI] Optional Features")
    $selection.TypeParagraph()
    $selection.TypeParagraph()

    $selection.Style = 'Normal'
    $selection.TypeText("This data is collected from the Win32_OptionalFeature WMI class. For more information please go to https://docs.microsoft.com/en-gb/windows/desktop/CIMWin32Prov/win32-optionalfeature")
    $selection.TypeParagraph()
    $selection.TypeParagraph()

        if(($xml.info.win32_optionalfeature).count -eq 0){
        Write-Output "[WARNING] WMI Optional Feature details not found. Moving on to next section..."
        $warnings ++
        $selection.Style = 'Normal'
        $selection.Font.Color="255"
        $selection.TypeText("[WARNING] WMI Optional Feature data not found!")
        $selection.TypeParagraph()
        }
        elseif(($xml.info.win32_optionalfeature.errorcode) -eq 1){
        Write-Output "[WARNING] WMI Optional Feature details not collected. Moving on to next section..."
        $warnings ++
        $selection.Style = 'Normal'
        $selection.Font.Color="255"
        $selection.TypeText("[WARNING] WMI Optional Feature data not collected!")
        $selection.TypeParagraph()
        $selection.TypeText("Reason for error: $($xml.info.win32_optionalfeature.errortext)")
        $selection.TypeParagraph()
        }
        else{
        Write-Output "[INFO] Populating WMI Optional Feature table."

        # Count rows in multi
        $multis = $xml.selectnodes("//info/win32_optionalfeature/win32_optionalfeature_multi")
        $count = ($multis | Measure-Object).Count

            # Calculate rows (from query above) and add 1 for the header
            if($count -eq 0){
            $rows = $count + 2
            }
            else{
            $rows = $count + 1
            }

        $table = $selection.Tables.add(
        $selection.Range,
        $rows,
        2,
        [Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior,
        [Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitWindow
        )

        $table.style = "Grid Table 4 - Accent 1"
        $table.cell(1,1).range.text = "Caption"
        $table.cell(1,2).range.text = "Name"

        $i = 2 # 2 as header is row 1

            if($count -eq 0){
            $table.cell($i,1).range.text = "No optional features installed"
            $table.Cell($i,2).Merge($table.Cell($i,1))
            }

            foreach($multi in $multis){

            $table.cell($i,1).range.text = $multi.caption
            $table.cell($i,2).range.text = $multi.name
            $i++

            }

        $table.Rows.item(1).Headingformat=-1
        $table.ApplyStyleFirstColumn = $false
        $selection.InsertCaption(-2, ": [WMI] Optional Features", $null, 1, $false)

        }

    $selection.EndOf(15) | Out-Null
    $selection.MoveDown() | Out-Null
    $selection.InsertNewPage()

        # Only include if Detailed is selected.
        if($reportType -eq "Detailed"){
        $selection.InsertNewPage()
        }
    #endregion ReportWMIOptionalFeatures

    #region ReportWMIServices
    # DETAILED
    if($reportType -eq "Detailed"){
    $selection.style = 'Heading 1'
    $selection.TypeText("[WMI] Services")
    $selection.TypeParagraph()
    $selection.TypeParagraph()

    $selection.Style = 'Normal'
    $selection.TypeText("This data is collected from the Win32_Service WMI class. For more information please go to https://docs.microsoft.com/en-gb/windows/desktop/CIMWin32Prov/win32-service")
    $selection.TypeParagraph()
    $selection.TypeParagraph()

        if(($xml.info.win32_service).count -eq 0){
        Write-Output "[WARNING] WMI Service details not found. Moving on to next section..."
        $warnings ++
        $selection.Style = 'Normal'
        $selection.Font.Color="255"
        $selection.TypeText("[WARNING] WMI Service data not found!")
        $selection.TypeParagraph()
        }
        elseif(($xml.info.win32_service.errorcode) -eq 1){
        Write-Output "[WARNING] WMI Service details not collected. Moving on to next section..."
        $warnings ++
        $selection.Style = 'Normal'
        $selection.Font.Color="255"
        $selection.TypeText("[WARNING] WMI Service data not collected!")
        $selection.TypeParagraph()
        $selection.TypeText("Reason for error: $($xml.info.win32_service.errortext)")
        $selection.TypeParagraph()
        }
        else{
        Write-Output "[INFO] Populating WMI Service table."

        # Count rows in multi
        $multis = $xml.selectnodes("//info/win32_service/win32_service_multi") | Where-Object {$_.source -ne 'System'} | Sort-Object {$_.name}
        $count = ($multis | Measure-Object).Count

        # Calculate rows (from query above) and add 1 for the header
        $rows = $count + 1

        $table = $selection.Tables.add(
        $selection.Range,
        $rows,
        4,
        [Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior,
        [Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitWindow
        )

        $table.style = "Grid Table 4 - Accent 1"
        $table.cell(1,1).range.text = "Name"
        $table.cell(1,2).range.text = "Start Mode"
        $table.cell(1,3).range.text = "Start Name"
        $table.cell(1,4).range.text = "State"

        $i = 2 # 2 as header is row 1

            foreach($multi in $multis){

            $table.cell($i,1).range.text = $multi.name
            $table.cell($i,2).range.text = $multi.startmode
            $table.cell($i,3).range.text = $multi.startname
            $table.cell($i,4).range.text = $multi.state
            $i++

            }

        $table.Rows.item(1).Headingformat=-1
        $table.ApplyStyleFirstColumn = $false
        $selection.InsertCaption(-2, ": [WMI] Services", $null, 1, $false)

        }

    $selection.EndOf(15) | Out-Null
    $selection.MoveDown() | Out-Null
    $selection.InsertNewPage()
    }
    #endregion ReportWMIServices

    #region ReportWMIProcesses
    # DETAILED
    if($reportType -eq "Detailed"){
    $selection.style = 'Heading 1'
    $selection.TypeText("[WMI] Processes")
    $selection.TypeParagraph()
    $selection.TypeParagraph()

    $selection.Style = 'Normal'
    $selection.TypeText("This data is collected from the Win32_Process WMI class. For more information please go to https://docs.microsoft.com/en-gb/windows/desktop/CIMWin32Prov/win32-process")
    $selection.TypeParagraph()
    $selection.TypeParagraph()

        if(($xml.info.win32_process).count -eq 0){
        Write-Output "[WARNING] WMI Process details not found. Moving on to next section..."
        $warnings ++
        $selection.Style = 'Normal'
        $selection.Font.Color="255"
        $selection.TypeText("[WARNING] WMI Process data not found!")
        $selection.TypeParagraph()
        }
        elseif(($xml.info.win32_process.errorcode) -eq 1){
        Write-Output "[WARNING] WMI Process details not collected. Moving on to next section..."
        $warnings ++
        $selection.Style = 'Normal'
        $selection.Font.Color="255"
        $selection.TypeText("[WARNING] WMI Process data not collected!")
        $selection.TypeParagraph()
        $selection.TypeText("Reason for error: $($xml.info.win32_process.errortext)")
        $selection.TypeParagraph()
        }
        else{
        Write-Output "[INFO] Populating WMI Process table."

        # Count rows in multi
        $multis = $xml.selectnodes("//info/win32_process/win32_process_multi")
        $count = ($multis | Measure-Object).Count

        # Calculate rows (from query above) and add 1 for the header
        $rows = $count + 1

        $table = $selection.Tables.add(
        $selection.Range,
        $rows,
        2,
        [Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior,
        [Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitWindow
        )

        $table.style = "Grid Table 4 - Accent 1"
        $table.cell(1,1).range.text = "Name"
        $table.cell(1,2).range.text = "Description"

        $i = 2 # 2 as header is row 1

            foreach($multi in $multis){

            $table.cell($i,1).range.text = $multi.name
            $table.cell($i,2).range.text = $multi.description
            $i++

            }

        $table.Rows.item(1).Headingformat=-1
        $table.ApplyStyleFirstColumn = $false
        $selection.InsertCaption(-2, ": [WMI] Processes", $null, 1, $false)

        }

    $selection.EndOf(15) | Out-Null
    $selection.MoveDown() | Out-Null
    $selection.InsertNewPage()
    }
    #endregion ReportWMIProcesses

    #region ReportWMISystemDrivers
    # DETAILED
    if($reportType -eq "Detailed"){
    $selection.style = 'Heading 1'
    $selection.TypeText("[WMI] System Drivers")
    $selection.TypeParagraph()
    $selection.TypeParagraph()

    $selection.Style = 'Normal'
    $selection.TypeText("This data is collected from the Win32_SystemDriver WMI class. For more information please go to https://docs.microsoft.com/en-gb/windows/desktop/CIMWin32Prov/win32-systemdriver")
    $selection.TypeParagraph()
    $selection.TypeParagraph()

        if(($xml.info.win32_systemdriver).count -eq 0){
        Write-Output "[WARNING] WMI System Driver details not found. Moving on to next section..."
        $warnings ++
        $selection.Style = 'Normal'
        $selection.Font.Color="255"
        $selection.TypeText("[WARNING] WMI System Driver data not found!")
        $selection.TypeParagraph()
        }
        elseif(($xml.info.win32_systemdriver.errorcode) -eq 1){
        Write-Output "[WARNING] WMI System Driver details not collected. Moving on to next section..."
        $warnings ++
        $selection.Style = 'Normal'
        $selection.Font.Color="255"
        $selection.TypeText("[WARNING] WMI System Driver data not collected!")
        $selection.TypeParagraph()
        $selection.TypeText("Reason for error: $($xml.info.win32_systemdriver.errortext)")
        $selection.TypeParagraph()
        }
        else{
        Write-Output "[INFO] Populating WMI System Driver table."

        # Count rows in multi
        $multis = $xml.selectnodes("//info/win32_systemdriver/win32_systemdriver_multi")
        $count = ($multis | Measure-Object).Count

        # Calculate rows (from query above) and add 1 for the header
        $rows = $count + 1

        $table = $selection.Tables.add(
        $selection.Range,
        $rows,
        5,
        [Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior,
        [Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitWindow
        )

        $table.style = "Grid Table 4 - Accent 1"
        $table.cell(1,1).range.text = "Name"
        $table.cell(1,2).range.text = "Service Type"
        $table.cell(1,3).range.text = "Start Mode"
        $table.cell(1,4).range.text = "Started"
        $table.cell(1,5).range.text = "State"

        $i = 2 # 2 as header is row 1

            foreach($multi in $multis){

            $table.cell($i,1).range.text = $multi.name
            $table.cell($i,2).range.text = $multi.servicetype
            $table.cell($i,3).range.text = $multi.startmode
            $table.cell($i,4).range.text = $multi.started
            $table.cell($i,5).range.text = $multi.state
            $i++

            }

        $table.Rows.item(1).Headingformat=-1
        $table.ApplyStyleFirstColumn = $false
        $selection.InsertCaption(-2, ": [WMI] System Drivers", $null, 1, $false)

        }

    $selection.EndOf(15) | Out-Null
    $selection.MoveDown() | Out-Null
    $selection.InsertNewPage()
    }
    #endregion ReportWMISystemDrivers

    #region ReportWMIPNPEntity
    # DETAILED
    if($reportType -eq "Detailed"){
    $selection.style = 'Heading 1'
    $selection.TypeText("[WMI] PNP Entities")
    $selection.TypeParagraph()
    $selection.TypeParagraph()

    $selection.Style = 'Normal'
    $selection.TypeText("This data is collected from the Win32_PNPEntity WMI class. For more information please go to https://docs.microsoft.com/en-gb/windows/desktop/CIMWin32Prov/win32-pnpentity")
    $selection.TypeParagraph()
    $selection.TypeParagraph()

        if(($xml.info.win32_pnpentity).count -eq 0){
        Write-Output "[WARNING] WMI PNP Entity details not found. Moving on to next section..."
        $warnings ++
        $selection.Style = 'Normal'
        $selection.Font.Color="255"
        $selection.TypeText("[WARNING] WMI PNP Entity data not found!")
        $selection.TypeParagraph()
        }
        elseif(($xml.info.win32_pnpentity.errorcode) -eq 1){
        Write-Output "[WARNING] WMI PNP Entity details not collected. Moving on to next section..."
        $warnings ++
        $selection.Style = 'Normal'
        $selection.Font.Color="255"
        $selection.TypeText("[WARNING] WMI PNP Entity Product data not collected!")
        $selection.TypeParagraph()
        $selection.TypeText("Reason for error: $($xml.info.win32_pnpentity.errortext)")
        $selection.TypeParagraph()
        }
        else{
        Write-Output "[INFO] Populating WMI PNP Entity table."

        # Count rows in multi
        $multis = $xml.selectnodes("//info/win32_pnpentity/win32_pnpentity_multi")
        $count = ($multis | Measure-Object).Count

        # Calculate rows (from query above) and add 1 for the header
        $rows = $count + 1

        $table = $selection.Tables.add(
        $selection.Range,
        $rows,
        4,
        [Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior,
        [Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitWindow
        )

        $table.style = "Grid Table 4 - Accent 1"
        $table.cell(1,1).range.text = "Name"
        $table.cell(1,2).range.text = "Class GUID"
        $table.cell(1,3).range.text = "Device ID"
        $table.cell(1,4).range.text = "Hardware ID"

        $i = 2 # 2 as header is row 1

            foreach($multi in $multis){

            $table.cell($i,1).range.text = $multi.name
            $table.cell($i,2).range.text = $multi.classguid
            $table.cell($i,3).range.text = $multi.deviceid
            $table.cell($i,4).range.text = $multi.hardwareid
            $i++

            }

        $table.Rows.item(1).Headingformat=-1
        $table.ApplyStyleFirstColumn = $false
        $selection.InsertCaption(-2, ": [WMI] PNP Entities", $null, 1, $false)

        }

    $selection.EndOf(15) | Out-Null
    $selection.MoveDown() | Out-Null
    $selection.InsertNewPage()
    }
    #endregion ReportWMIPNPEntity

    #region ReportCertificates
    # BASIC & DETAILED
    $selection.style = 'Heading 1'
    $selection.TypeText("[System] Computer Certificates")
    $selection.TypeParagraph()
    $selection.TypeParagraph()

    $selection.Style = 'Normal'
    $selection.TypeText("This data is collected from the System.Security.Cryptography.X509Certificates.X509Store. For more information please go to https://docs.microsoft.com/en-us/dotnet/api/system.security.cryptography.x509certificates.x509store?view=net-5.0")
    $selection.TypeParagraph()
    $selection.TypeParagraph()


        if($reportType -ne 'Basic'){
        $selection.style = 'Heading 2'
        $selection.TypeText("Basic")
        $selection.TypeParagraph()
        $selection.TypeParagraph()
        }

        if(($xml.info.computer_certificates).count -eq 0){
        Write-Output "[WARNING] System Computer Certificte (Basic) details not found. Moving on to next section..."
        $warnings ++
        $selection.Style = 'Normal'
        $selection.Font.Color="255"
        $selection.TypeText("[WARNING] System Computer Certificate (Basic) data not found!")
        $selection.TypeParagraph()
        }
        elseif(($xml.info.computer_certificates.errorcode) -eq 1){
        Write-Output "[WARNING] System Computer Certificate (Basic) details not collected. Moving on to next section..."
        $warnings ++
        $selection.Style = 'Normal'
        $selection.Font.Color="255"
        $selection.TypeText("[WARNING] System Computer Certificate (Basic) data not collected!")
        $selection.TypeParagraph()
        $selection.TypeText("Reason for error: $($xml.info.computer_certificates.errortext)")
        $selection.TypeParagraph()
        }
        else{
        Write-Output "[INFO] Populating System Computer Certificate (Basic) table."

        # Count rows in multi
        $multis = $xml.selectnodes("//info/computer_certificates/computer_certificates_multi")
        $count = ($multis | Measure-Object).Count

            # Calculate rows (from query above) and add 1 for the header
            if($count -eq 0){
            $rows = $count + 2
            }
            else{
            $rows = $count + 1
            }

        $table = $selection.Tables.add(
        $selection.Range,
        $rows,
        5,
        [Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior,
        [Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitWindow
        )

        $table.style = "Grid Table 4 - Accent 1"
        $table.cell(1,1).range.text = "Store"
        $table.cell(1,2).range.text = "Subject"
        $table.cell(1,3).range.text = "Issuer"
        $table.cell(1,4).range.text = "Valid From"
        $table.cell(1,5).range.text = "Expiration"

        $i = 2 # 2 as header is row 1

            foreach($multi in $multis){

            $table.cell($i,1).range.text = $multi.store
            $table.cell($i,2).range.text = $multi.subject
            $table.cell($i,3).range.text = $multi.issuer
            $table.cell($i,4).range.text = $multi.validfrom
            $table.cell($i,5).range.text = $multi.expiration
            $i++

            }

        $table.Rows.item(1).Headingformat=-1
        $table.ApplyStyleFirstColumn = $false
        $selection.InsertCaption(-2, ": [System] Computer Certificates (Basic)", $null, 1, $false)

          if($reportType -ne "Basic"){
          $selection.TypeParagraph()
          }

        }

    $selection.EndOf(15) | Out-Null
    $selection.MoveDown() | Out-Null

        if($reportType -eq "Detailed"){
        $selection.TypeParagraph()
        $selection.style = 'Heading 2'
        $selection.TypeText("Detailed")
        $selection.TypeParagraph()
        $selection.TypeParagraph()

            if(($xml.info.computer_certificates).count -eq 0){
            Write-Output "[WARNING] System Computer Certificates (Detailed) details not found. Moving on to next section..."
            $warnings ++
            $selection.Style = 'Normal'
            $selection.Font.Color="255"
            $selection.TypeText("[WARNING] System Computer Certificates (Detailed) data not found!")
            $selection.TypeParagraph()
            }
            elseif(($xml.info.computer_certificates.errorcode) -eq 1){
            Write-Output "[WARNING] System Computer Certificates (Detailed) details not collected. Moving on to next section..."
            $warnings ++
            $selection.Style = 'Normal'
            $selection.Font.Color="255"
            $selection.TypeText("[WARNING] System Computer Certificates (Detailed) data not collected!")
            $selection.TypeParagraph()
            $selection.TypeText("Reason for error: $($xml.info.computer_certificates.errortext)")
            $selection.TypeParagraph()
            }
            else{
            Write-Output "[INFO] Populating System Computer Certificates (Detailed) table."

            # Count rows in multi
            $multis = $xml.selectnodes("//info/computer_certificates/computer_certificates_multi")
            $count = ($multis | Measure-Object).Count

                # Calculate rows. (No of items x no of multis) + (no of multis + 1 (for header)) | or less 1 for header if single row
                if($count -eq 0){
                $rows = $count + 2
                }
                elseif($count -eq 1){
                $rows = (15 + $count)
                }
                else{
                $rows = (15 * $count) + ($count + 1)
                }

            $table = $selection.Tables.add(
            $selection.Range,
            $rows,
            2,
            [Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior,
            [Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitWindow
            )

            $table.style = "Grid Table 4 - Accent 1"
            $table.cell(1,1).range.text = "Item"
            $table.cell(1,2).range.text = "Value"

            $i = 2 # 2 as header is row 1
            $y = 1

                foreach($multi in $multis){

                    if($count -gt 1){
                    $table.cell($i,1).Merge($table.cell($i, 2))
                    $table.cell($i,1).range.text = "------ BLOCK $y ------"
                    $i++
                    }

                $table.cell($i,1).range.text = "Store"
                $table.cell($i,2).range.text = $multi.store
                $i++

                $table.cell($i,1).range.text = "Subject"
                $table.cell($i,2).range.text = $multi.subject
                $i++

                $table.cell($i,1).range.text = "Issuer"
                $table.cell($i,2).range.text = $multi.issuer
                $i++

                $table.cell($i,1).range.text = "Valid From"
                $table.cell($i,2).range.text = $multi.validfrom
                $i++

                $table.cell($i,1).range.text = "Expiration"
                $table.cell($i,2).range.text = $multi.expiration
                $i++

                $table.cell($i,1).range.text = "Thumbprint"
                $table.cell($i,2).range.text = $multi.thumbprint
                $i++

                $table.cell($i,1).range.text = "Serial Number"
                $table.cell($i,2).range.text = $multi.serialnumber
                $i++

                $table.cell($i,1).range.text = "Format"
                $table.cell($i,2).range.text = $multi.format
                $i++

                $table.cell($i,1).range.text = "Version"
                $table.cell($i,2).range.text = $multi.version
                $i++

                $table.cell($i,1).range.text = "Signature Algorithm Friendly Name"
                $table.cell($i,2).range.text = $multi.signaturealgorithmfriendlyname
                $i++

                $table.cell($i,1).range.text = "Signature Algorithm Value"
                $table.cell($i,2).range.text = $multi.signaturealgorithmvalue
                $i++

                $table.cell($i,1).range.text = "Enhanced Key Usage List Friendly Name"
                $table.cell($i,2).range.text = $multi.enhancedkeyusagelistfriendlyname
                $i++

                $table.cell($i,1).range.text = "Archived"
                $table.cell($i,2).range.text = $multi.archived
                $i++

                $table.cell($i,1).range.text = "Has Private Key"
                $table.cell($i,2).range.text = $multi.hasprivatekey
                $i++

                $table.cell($i,1).range.text = "Friendly Name"
                $table.cell($i,2).range.text = $multi.friendlyname
                $i++

                $y++

                }

            $table.Rows.item(1).Headingformat=-1
            $table.ApplyStyleFirstColumn = $false
            $selection.InsertCaption(-2, ": [System] Computer Certificates (Detailed)", $null, 1, $false)

            }

        }

    $selection.EndOf(15) | Out-Null
    $selection.MoveDown() | Out-Null
    $selection.InsertNewPage()
    #endregion ReportCertificates

    #region ReportWMIPowerPlan
    # BASIC
    $selection.style = 'Heading 1'
    $selection.TypeText("[WMI] Power Plan")
    $selection.TypeParagraph()
    $selection.TypeParagraph()

    $selection.Style = 'Normal'
    $selection.TypeText("This data is collected from the Win32_PowerPlan WMI class. For more information please go to https://docs.microsoft.com/en-us/previous-versions/windows/desktop/legacy/dd904531(v=vs.85)")
    $selection.TypeParagraph()
    $selection.Style = 'Normal'

        if(($xml.info.win32_powerplan).count -eq 0){
        Write-Output "[WARNING] WMI Power Plan details not found. Moving on to next section..."
        $warnings ++
        $selection.Style = 'Normal'
        $selection.Font.Color="255"
        $selection.TypeText("[WARNING] WMI Power Plan data not found!")
        $selection.TypeParagraph()
        }
        elseif(($xml.info.win32_powerplan.errorcode) -eq 1){
        Write-Output "[WARNING] WMI Power Plan details not collected. Moving on to next section..."
        $warnings ++
        $selection.Style = 'Normal'
        $selection.Font.Color="255"
        $selection.TypeText("[WARNING] WMI Power data not collected!")
        $selection.TypeParagraph()
        $selection.TypeText("Reason for error: $($xml.info.win32_powerplan.errortext)")
        $selection.TypeParagraph()
        }
        else{
        Write-Output "[INFO] Populating WMI Power Plan table."

        # Count rows in multi
        $multis = $xml.selectnodes("//info/win32_powerplan/win32_powerplan_multi")
        $count = ($multis | Measure-Object).Count

            # Calculate rows. (No of items x no of multis) + (no of multis + 1 (for header)) | or less 1 for header if single row
            if($count -eq 1){
            $rows = (4 + $count)
            }
            else{
            $rows = (4 * $count) + ($count + 1)
            }

        $table = $selection.Tables.add(
        $selection.Range,
        $rows,
        2,
        [Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior,
        [Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitWindow
        )

        $table.style = "Grid Table 4 - Accent 1"
        $table.cell(1,1).range.text = "Item"
        $table.cell(1,2).range.text = "Value"

        $i = 2 # 2 as header is row 1
        $y = 1

            foreach($multi in $multis){

                if($count -gt 1){
                $table.cell($i,1).Merge($table.cell($i, 2))
                $table.cell($i,1).range.text = "------ BLOCK $y ------"
                $i++
                }

            $table.cell($i,1).range.text = "Element Name"
            $table.cell($i,2).range.text = $multi.elementname

            $i++

            $table.cell($i,1).range.text = "Description"
            $table.cell($i,2).range.text = $multi.description
            $i++

            $table.cell($i,1).range.text = "Instance ID"
            $table.cell($i,2).range.text = $multi.instanceid
            $i++

            $table.cell($i,1).range.text = "Is Active"
            $table.cell($i,2).range.text = $multi.isactive

            $i++

            $y++

            }

        $table.Rows.item(1).Headingformat=-1
        $table.ApplyStyleFirstColumn = $false
        $selection.InsertCaption(-2, ": [WMI] Power Plan", $null, 1, $false)

        }

    $selection.EndOf(15) | Out-Null
    $selection.MoveDown() | Out-Null
    $selection.InsertNewPage()
    #endregion ReportWMIPowerPlan

    #region ReportFinalise
    ### UPDATE TABLE OF CONTENTS ###
    $toc.update()

    # Rename .xml to .docx and add report type
    if(($mode -eq "GatherAndReport") -or ($mode -eq "GatherOnly")){
    $reportType = (Get-Culture).TextInfo.ToTitleCase($reportType)
    $reportName = $reportFile -replace ".xml", ".docx"
    $reportName = $reportName -replace ".docx", "_$reportType.docx"
    }
    # Rename .xml to .docx and add report type. Get file name and use output directory to construct path to save file to
    else{
    $reportType = (Get-Culture).TextInfo.ToTitleCase($reportType)
    $reportName = Split-Path $reportFile -Leaf
    $reportName = "$outDir\$reportName"
    $reportName = $reportName -replace ".xml", ".docx"
    $reportName = $reportName -replace ".docx", "_$reportType.docx"
    }

        try{
        Write-Output "[INFO] Saving $reportName."
        $document.SaveAs([ref]$reportName,[ref]$saveFormat::wdFormatDocument)
        }
        catch{
        Write-Output "[ERROR] Unable to save $reportName. Script terminated!"
        Write-Output "[ERROR] $($_.exception.message)"
        break
        }


    $word.Quit()

    $null = [System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$word)
    [gc]::Collect()
    [gc]::WaitForPendingFinalizers()

    } # end -ne GatherOnly
    #endregion ReportFinalise

#endregion Report

#region CompleteMessage
    ### DISPLAY COMPLETE MESSAGE ###
    Write-Output ""

    if($warnings -eq 0){
    Write-Output "[INFO] Script complete with $warnings warnings, in $($sw.Elapsed.Hours) hours, $($sw.Elapsed.Minutes) minutes, $($sw.Elapsed.Seconds) seconds."
    }
    elseif($warnings -eq 1){
    Write-Output "[WARNING] Script complete with $warnings warning, in $($sw.Elapsed.Hours) hours, $($sw.Elapsed.Minutes) minutes, $($sw.Elapsed.Seconds) seconds."
    }
    else{
    Write-Output "[WARNING] Script complete with $warnings warnings, in $($sw.Elapsed.Hours) hours, $($sw.Elapsed.Minutes) minutes, $($sw.Elapsed.Seconds) seconds."
    }

Write-Output ""
#endregion CompleteMessage

#region Comments
<#

### REPORT SECTION ###

The following sections will report no data comments in the table, if applicable. All other items are deemed to always have data in them

- [Registry] Installed Programs (System)
    - Basic
    - Advanced
- [Registry] Installed Programs (User)
    - Basic
    - Advanced
- [WMI] Share
- [WMI] Start Up Command
    - Basic
    - Detailed
- [WMI] Page File Usage
- [WMI] Quick Fix Engineering (Updates)
    - Basic
    - Detailed
- [WMI] CD ROM Drive
- [WMI] Video Controller
- [WMI] Sound Device
- [WMI] Printers
   - Basic
   - Detailed
 - [WMI] Encryptable Volume
 - [WMI] Firewall Product
 - [WMI] Antivirus Product
 - [WMI] Antispyware Product
 - [WMI] Optional Features

#>
#endregion Comments
