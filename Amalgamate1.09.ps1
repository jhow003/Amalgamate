#Amalgamate
#v 1.09
#author:Jeff Howerton

#This script is for all whom have suffered thru the pain of merging excel docs and removing duplicate enteries. (i.e. SANS students)


function Write-ProgressHelper {
	param (
	    [int]$StepNumber,
	    [string]$Message
	)

	Write-Progress -Activity 'Title' -Status $Message -PercentComplete (($StepNumber / $steps) * 100)
}

$script:steps = ([System.Management.Automation.PsParser]::Tokenize((gc "$PSScriptRoot\$($MyInvocation.MyCommand.Name)"), [ref]$null) | where { $_.Type -eq 'Command' -and $_.Content -eq 'Write-ProgressHelper' }).Count

$stepCounter = 0

Write-ProgressHelper -Message 'Collecting Data' -StepNumber ($stepCounter++)
Start-Sleep -Seconds 5

#import from excel doc, add parameters/headers
$imported1 = ($x = Import-Excel -Path .\test.xlsx  -HeaderName 'KEYWORD', 'BOOK', 'PAGE', 'DEFINITION') 

#Write-Progress -Activity "Importing Data" -Id 1 -Status "Processing $i/$($imported.count) imports" -PercentComplete ($i/ $imported.count *100)


#sort first column, filter duplicates
$sorted1 = $imported1 | Sort-Object -Property @{Expression = "KEYWORD"} | Get-Unique -AsString 


#produce new excel doc
$sorted1 | Export-excel -Path .\output1.xlsx 

# 512_Raw_combined


#second pass
Write-ProgressHelper -Message 'Filtering Second Pass' -StepNumber ($stepCounter++)

Start-Sleep -Seconds 5

#import from excel doc, add parameters/headers
$imported2 = ($x = Import-Excel -Path .\output1.xlsx) 

#sort fourth column, filter duplicates
$sorted2 = $imported2 | Sort-Object -Property @{Expression = "DEFINITION"} | Get-Unique -AsString 

#produce new excel doc 
$sorted2 | Export-excel -Path .\output2.xlsx 



#3rdpass
Write-ProgressHelper -Message 'Filtering Third Pass' -StepNumber ($stepCounter++)

Start-Sleep -Seconds 5

#import from excel doc
$imported3 = ($x = Import-Excel -Path .\output2.xlsx)

#sort first column, filter duplicates
$sorted3 = $imported3 | Sort-Object -Property @{Expression = "KEYWORD"} | Get-Unique -AsString 


#produce new excel doc and format then open it for you when its finished
$sorted3 | Export-excel -Path .\output3.xlsx 





#final pass
Write-ProgressHelper -Message 'Finishing Up' -StepNumber ($stepCounter++)

Start-Sleep -Seconds 5

#import from excel doc
$imported4 = ($x = Import-Excel -Path .\output3.xlsx) 

#sort fourth column, filter duplicates
$sorted4 = $imported4 | Sort-Object -Property @{Expression = "DEFINITION"} | Get-Unique -AsString 

#final sort
$sorted4 = $imported4 | Sort-Object -Property @{Expression = "KEYWORD"}

#produce new excel doc and format then open it for you when its finished
$sorted4 | Export-excel -Path .\FINAL$((get-date).tostring("MM-dd-yyyy.T.HH-mm-ss")).xlsx -BoldTopRow -AutoFilter -FreezeTopRow -AutoSize -Show

#clean up
rm .\output1.xlsx
rm .\output2.xlsx
rm .\output3.xlsx

#NOTES# 
# will not retain text formats at this time
#ADD help file data
