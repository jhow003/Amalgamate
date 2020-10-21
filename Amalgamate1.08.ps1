#name of script TBD
#v 1.08
#author:Jeff Howerton

#This script is for all whom have suffered thru the pain of merging excel docs and removing duplicate enteries. (i.e. SANS students)


#import from excel doc, add parameters/headers
$imported = ($x = Import-Excel -Path .\512_Raw_combined.xlsx  -HeaderName 'KEYWORD', 'BOOK', 'PAGE', 'DEFINITION') 
#sort first column, filter duplicates
$sorted = $imported | Sort-Object -Property @{Expression = "KEYWORD"} | Get-Unique -AsString 
#produce new excel doc and format then open it for you when its finished
$sorted | Export-excel -Path .\output.xlsx -BoldTopRow -AutoFilter -FreezeTopRow -AutoSize -Show

#NOTES# will not retain text formats at this time
#ADD help file data
#ADD usage README
#ADD status bar for script progress