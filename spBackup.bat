@echo off 
cd C:\Users\admin\Desktop\spbackup
PowerShell.exe -command ".\cvsImportToSQL.ps1" "172.18.0.12\spbackup" "splist" "dbo.sptable" 





