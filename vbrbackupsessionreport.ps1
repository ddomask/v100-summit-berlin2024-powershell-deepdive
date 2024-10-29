<#
  .SYNOPSIS
  
    Generate a report on Backup Sessions based on the time scope defined in the Parameters. 
	Use -TimeScope or -StartTime and -EndTime to define a time scope.
	Use -ExportAs to export to CSV or HTML. You must define a path to export the report for with -ReportExportPath when using -ExportAs
	If you do not set -ExportAs, the results wil be printed to the powershell console.
	
	.\vbrbackupsessionreport.ps1 -TimeScope [Last 24 Hours, Last 7 Days, All Time] -ExportAs [CSV, HTML] -ReportExportPath [string]
	.\vbrbackupsessionreport.ps1 -StartTime [DateTime Object] -EndTime [DateTime Object] -ExportAs [CSV, HTML] -ReportExportPath [string]


	.DESCRIPTION
	  Returns Backup Sessions from a given time range and allows you to export to shell, CSV ,or HTML.

	.PARAMETER TimeScope
	  Set the scope of time you'd like to generate a session report for. Options are Last 24 Hours, Last 7 Days, All Time. 
	  To set a custom time scope, use the -StartTime and -EndTime parameters and do not use -TimeScope

	.PARAMETER ExportAs
	  Script will generate a report for the given time scope and then export it to either CSV or HTML.
	  
	.PARAMETER StartTime
	  Accepts a DateTime object returned by Get-Date. Sets the beginning of the time scope you'd like to export from.  
	  
	.PARAMETER EndTime
	  Accepts a DateTime object returned by Get-Date. Sets the ending of the time scope you'd like to export from.
	  
	.PARAMETER ReportExportPath
	  When exporting as CSV or HTML, define where to write the report file.

	.INPUTS
	  None.

	.OUTPUTS
	  Depends

	.EXAMPLE
	  C:\PS> .\vbrbackupsessionreport.ps1 -TimeScope 'Last 24 Hours' 
	  
	  Export backup sessions from the last 24 Hours to the shell

	.EXAMPLE
	  C:\PS> .\vbrbackupsessionreport.ps1 -TimeScope 'Last 7 Days' -ExportAs CSV -ReportExportPath C:\temp
	  
	  Export backup sessions from the last 7 days and save it to a .CSV file in C:\temp. 

	.EXAMPLE
	  C:\PS> .\vbrbackupsessionreport.ps1 -StartTime $start -EndTime $end 
	  Export backup sessions between the $start date and the $end date. $start and $end were created with $start = Get-Date -Date 18.4.1816; $end = Get-Date -Date 1.4.2024. You can enter string values as well for dates (e.g., for October 31st 2024, 31.10.2024 will be accepted. Use the date/time format you are used to it will respect the system time format.
#>
using namespace System.Collections.Generic

param(
	[Parameter(Mandatory = $false)]
		[ValidateSet("Last 24 Hours","Last 7 days","All Time")]
		[string]$TimeScope,
	[Parameter(Mandatory = $false)]
		[ValidateSet('CSV','HTML')]
		$ExportAs,
	[Parameter(Mandatory = $false)]
		[DateTime]$StartTime,
	[Parameter(Mandatory = $false)]
		[DateTime]$EndTime,
	[Parameter(Mandatory = $false)]
		[string]$ReportExportPath
)

#User Message Strings
#These are placed out of the way to keep the core part of the script lean. We pass it to the word-wrap function so we don't have to format stuff.

$NoRangeError = "No Time range was selected for reporting or a Start/End time was not provided. Use the -TimeScope parameter to export last 24 hours, last 7 days, or All time. Use -StartTime and -EndTime to enter a custom time range. Use 'Get-Help $($MyInvocation.MyCommand)' to see the script README, or simply run the script without any parameters"

$NoExportPathError = "No path for the report export was set. When using -ExportAs, you must define a path for the export. Use the -ReportExportPath parameter and define a path where the file should be placed. For example, if you wish to write an HTML report to C:\temp, enter 'C:\temp' for the -ReportExportPath parameter"

$StartingSessionReporting = "Starting report generation. The script may not show any output while the data is collected from Veeam Backup and Replication, please allow it some time to process."
	
## Probably some functions here	

function Test-ExportPathExists {
	If(-Not(Test-Path $ReportExportPath)){
	Try{
		Write-Host "Export path for reports does not exist. Creating path $($ReportExportPath)"
		New-Item -ItemType Directory -Force -Path $ReportExportPath | Out-Null
	} catch {
		Write-Host -ForegroundColor Red "Could not create path $($ReportExportPath). Using C:\tmp as Report Export Path"
		$ReportExportPath = "C:\tmp"
		}
	}
}
	
function Get-SessionsWithDateRange {
	$Date = Get-Date
	switch ($TimeScope) {
		{$_ -like "*24*"} {$StartTime = $Date.AddDays(-1);$EndTime = $Date}
		{$_ -like "*7*"} {$StartTime = $Date.AddDays(-7);$EndTime = $Date}
		{$_ -Like "All*"} {$StartTime = $Date.AddDays(-100000);$EndTime = $Date}
	}
	$sortedSess = $fullSess | Where-Object {$_.CreationTime -ge $StartTime -and $_.CreationTime -le $EndTime}
	return $sortedSess
}

function word-wrap { #Created by StackOverflow user rojo: https://stackoverflow.com/a/35134216
    [CmdletBinding()]
    Param(
        [parameter(Mandatory=1,ValueFromPipeline=1,ValueFromPipelineByPropertyName=1)]
        [Object[]]$chunk
    )
    PROCESS {
        $Lines = @()
        foreach ($line in $chunk) {
            $str = ''
            $counter = 0
            $line -split '\s+' | %{
                $counter += $_.Length + 1
                if ($counter -gt $Host.UI.RawUI.BufferSize.Width) {
                    $Lines += ,$str.trim()
                    $str = ''
                    $counter = $_.Length + 1
                }
                $str = "$str$_ "
            }
            $Lines += ,$str.trim()
        }
        $Lines
    }
}

##Pre-run failure checks

If($PSBoundParameters.Values.Count -eq 0 -and $args.count -eq 0) {
    Get-Help $MyInvocation.MyCommand.Definition -Examples
    Break	
}elseif(-not($TimeScope) -And (-not($StartTime) -or -not($EndTime))){
	$NoRangeError | word-wrap | Write-Host -ForegroundColor Yellow
	Break
}

If($ExportAs -AND -not($ReportExportPath)){
	$NoExportPathError | word-wrap | Write-Host -ForegroundColor Yellow
	Break
}

	
	
##Core logic
$StartingSessionReporting | word-wrap | Write-Host -ForegroundColor Green
$backups = Get-VBRBackup
$jobTypes = $backups.JobType | Sort-Object -Unique
$sessList = [List[Object]]@()

Foreach($jt in $jobTypes){
	$tempSess = Get-VBRSession -Type $jt
	$sessList.Add($tempSess)
}

$fullSess = Get-VBRBackupSession -Id $sessList.id | Sort-Object -Property CreationTime -Descending
$realSessData = Get-SessionsWithDateRange

$sessPrep = [List[Object]]@()
Foreach($rs in $realSessData){
	$dataObject = [PSCustomObject]@{ #We build a PSCustomObject to simplify editing later on. We select from the array $sesPrep in the HTMLREPORTBUILD section and CSVREPORTBUILD section
		Name = $rs.Name
		JobType = $rs.JobType
		StartTime = $rs.CreationTime
		EndTime = $rs.EndTime
		Duration = ($rs.EndTime - $rs.CreationTime).ToString().Split(".")[0]
		Result = $rs.Result
	}
$sessPrep.Add($dataObject)
}


If($ExportAs -eq "HTML"){
#HTMLREPORTBUILD

	If(-not($TimeScope)){
		$ReportTime = "$($StartTime.ToShortDateString()) - $($EndTime.ToShortDateString())"
	} else {
		$ReportTime = $TimeScope
	}
	$Header = @"
<style>
	TABLE {border-width: 1px; border-style: solid; border-color: black; border-collapse: collapse;}
	TD {border-width: 1px; padding: 1px; border-style: solid; border-color: black; font-family: sans-serif}

	.resultWarning {
		background-color: #fff200;
	}
	.resultFailed {
		background-color: #ff0000;
	}
	
</style>
"@

	$sessHTMLPrep = $sessPrep | Select-Object Name, JobType, StartTime, EndTime, Duration, Result | ConvertTo-HTML -As Table -Fragment | Out-String 
	$reportHTML = ConvertTo-HTML -PostContent $sessHTMLPrep -PreContent "<h1>Veeam Backup Session Report - $($ReportTime)</h1>" -Head $Header 
	$reportHTML = $reportHTML -replace "<td>Warning</td>",'<td class="resultWarning">Warning</td>'
	$reportHTML = $reportHTML -replace "<td>Failed</td>",'<td class="resultFailed">Failed</td>'
	Test-ExportPathExists
	$reportHTML | Out-File $ReportExportPath\SessionsReport.html
	Write-Host "HTML Report Exported to $($ReportExportPath)"
	Start $ReportExportPath
	Break
} elseif($ExportAs -eq "CSV"){
#CSVREPORTBUILD
	Test-ExportPathExists
	$sessCSVPrep = $sessPrep | Select-Object Name, JobType, StartTime, EndTime, Duration, Result | Export-CSV -NoTypeInformation -Path $ReportExportPath\SessionsReport.csv
	Start $ReportExportPath
} else {
	return $realSessData 
} 


