Remove-Item -Path "$PSScriptRoot\output\*" 
Remove-Item -Path "$PSScriptRoot\input\*" 
Copy-Item -Path "$PSScriptRoot\testfile.xls" "$PSScriptRoot\input\"
