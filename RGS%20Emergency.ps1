<#	
	.NOTES
	===========================================================================
	 Created on:   	04/11/2014 8:13 p.m.
	 Created by:   	Paul Bloem
	 Blog: 	UCSorted.com
	 Twitter twitter.com/PaulB_NZ
	 Filename:     	
	===========================================================================
	.DESCRIPTION
		This script takes a DDI assigned to a RGS Workflow and re-assigns it to a Emergency Workflow - and vice versa depending on the 
		menu selection. You will need a pre-configured 
		Emergency RGS Queue, Group and Workflow. For simplicity you can configure the Emergency Queue to overflow to a destination with
		a Threshold of 0 (meaning that any calls will overflow immediatly.
		You then have the choice on the Emergency Queue to deliver the calls to another SIP address, a PSTN number (such as a cell phone)
		or to a UM enabled mailbox (Voicemail).
#>
# Script Config
Import-Module Lync
#The following Variables need to be manually edited to match your environment
$RGSWorkflow = "Support Workflow"									#RGS Workflow you want to failover
$RGSWorkflowDDI = "tel:+123456789"									#Line URI of the RGS Workflow
$EmergencyWorkflow = "Emergency - Support Workflow"					#Name of the Emergency Workflow
$sipdomain = "@ucsorted.com" 										#Lync Sip Domain
$ServiceID = "service:ApplicationServer:lyncFEPool1.ucsorted.com" 	#Lync front end pool identity
$logfile = "c:\RGSEmergency-log.txt"								#Specify a log file for this script

# Menu
#The Menu items below are displayed as TXT for selection purpose only. Once selected the next section (Selection Logic) is engaged
cls

Write-Host "* Set Emergency state for" $RGSWorkflow "" -ForegroundColor Cyan
Write-Host " "
Write-Host "Select one of the following options (Y\N):"
Write-Host " "
Write-Host "Y - Activate Emergency RGS" -foregroundcolor Yellow
Write-Host "N - De-Activate Emergency RGS" -foregroundcolor Yellow
Write-Host " "
$a = Read-Host "Select Y or N: "
Write-Host " "

# Selection Logic to set DDI for RGS based on Menu Selection

switch ($a)
{
	Y {
		$ActiveWorkflow = $EmergencyWorkflow
		$DeactiveWorkflow = $RGSWorkflow
	}
	N {
		$ActiveWorkflow = $RGSWorkflow
		$DeactiveWorkflow = $EmergencyWorkflow
	}
}

#Removing DDI from Active RGS
$a = Get-CsRgsWorkflow -Identity $ServiceID -name $DeactiveWorkflow
$a.LineUri = ""
set-csrgsworkflow -Instance $a

#Assigning the DDI to new Active RGS
$b = Get-CsRgsWorkflow -Identity $ServiceID -name $Activeworkflow
$b.LineUri = $RGSWorkflowDDI
set-csrgsworkflow -Instance $b

# Write to log file
$date = Get-Date
Write-output "$date, $ActiveWorkFlow, $RGSWorkflow1DDI" | Out-File -append $logfile

Write-Host "Active RGS changed to: " -NoNewline
Write-Host $ActiveWorkFlow -ForegroundColor Green
Write-Host " "
