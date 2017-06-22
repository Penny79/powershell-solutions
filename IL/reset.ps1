Remove-Item -Path "$PSScriptRoot\output\*" 
Remove-Item -Path "$PSScriptRoot\input\*" 
Copy-Item -Path "$PSScriptRoot\201702_vom_20170130_Fahrplan_2017_01 (2).xls" "$PSScriptRoot\input\"
