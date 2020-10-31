# Get-Win10Info

This PowerShell script gathers information from either a local or remote Windows 10 computer and produces a word report with that information.

## Modes
This script has 3 modes
- GatherAndReport
- GatherOnly
- ReportOnly

#### GatherAndReport

This mode will gather information from a Windows 10 computer and generate a word document from that information. This is the default mode so you don't need to specify this mode when running the script. This mode can also run with both the -endpoint and -reportMode parameters. 
###### -endpoint

This is used to target a remote computer. The remote computer will need to be configured to allow remote WMI queries and allow remote reading of the registry for this to work.
###### -reportMode

There are 2 report modes Basic and Detailed. These modes determine what level of information is contained within the word report.

#### GatherOnly

This mode will only gather information from a Windows 10 computer but will not generate the word report. To use this mode specify -mode ReportOnly when running this script. This mode can also run with the -endpoint parameter as mentioned above.

#### ReportOnly

This mode will generate a word report from a specified xml file that was created in either GatherAndReport or GatherOnly mode. When using this mode the -xmlReport parameter must be used. When using -xmlReport specify the location of the xml file created previously. This mode can also be used with the -reportMode parameter as mentioned above.

## Parameters

.PARAMETER outDir
This is the directory where the xml files and reports are stored. If the directory doesn't exist it will be created.

.PARAMETER mode
[OPTIONAL] This is the mode the script runs in. There are 3 modes available. GatherAndReport, GatherOnly & ReportOnly. GatherAndReport will collect details and generate a word document. GatherOnly will collect details only. ReportOnly will construct a word document from a provided xml file with previously gathered details. The default mode is GatherAndReport.

.PARAMETER endpoint
[OPTIONAL] This is the name of the remote endpoint if the script is to gather information from a remote endpoint. If this parameter is not specified then the script will run against the endpoint that the script is running on.

.PARAMETER xmlReport
[OPTIONAL] This parameter is only used for ReportOnly mode and is the xml file with the details previously gathered using GatherAndReport or GatherOnly mode.

.PARAMETER reportMode
[OPTIONAL] This is the report mode of the script. It can either be Basic or Detailed. Basic will report a smaller set of data whilst Detailed will provide a greater level of report.

## Examples

.EXAMPLE
.\Get-Win10Info.ps1 -outDir "c:\temp"
Runs the script in GatherAndReport mode which will save a basic report to c:\temp.

.EXAMPLE
.\Get-Win10Info.ps1 -outDir "c:\temp" -endpoint PC01
Runs the script in GatherAndReport mode which will save a basic report to c:\temp with the details gather from the remote endpoint PC01.

.\Get-Win10Info.ps1 -outDir "c:\temp" -mode GatherOnly -endpoint PC01
Runs the script in GatherOnly and collects details from the remote endpoint PC01.

.\Get-Win10Info.ps1 -outDir "c:\temp" -mode ReportOnly -xmlFile c:\temp\PC01_201909291709.xml -reportType Detailed
Runs the script in ReportOnly mode and generates a detailed report using GR001_201909291709.xml.

## Why This Script

There are many scripts that do this kind of thing but I don't get to write many scripts nowadays so I thought I would try to keep my hand in and write this one.

## Script Info

All the commands in the Gather portion of the script need to support getting information from a remote computer natively. I have used Get-WMIObject over Get-CIMInstance as over the years I have found Get-WMIObject to be more robust over Get-CIMInstance.

## Future Updates

- Bulk Gather 
- Bulk Report
- Gather more information.

## Feedback

Please use GitHub Issues to report any, well.... issues with the script.
