New-Item C:\Windows\System32\WindowsPowerShell\v1.0\Modules\SysAdmin -ItemType Directory
Copy-item '\\newton\admin\SysAdmin Powershell Module\SysAdmin.psd1','\\newton\admin\SysAdmin Powershell Module\SysAdmin.psm1' -Destination C:\windows\System32\WindowsPowerShell\v1.0\Modules\Sysadmin\
Import-Module C:\Windows\System32\WindowsPowerShell\v1.0\Modules\SysAdmin\SysAdmin.psd1 -Force

