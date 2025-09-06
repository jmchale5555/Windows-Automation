$remoteComputers = @(
"tlab-r11pc1",
"tlab-r11pc2",
"tlab-r11pc3",
"tlab-r11pc4",
"tlab-r11pc5"
)

foreach ($computer in $remoteComputers)
{
 $scriptblock = { 
       c:\windows\System32\WindowsPowerShell\v1.0\powershell.exe -ExecutionPolicy ByPass -NoExit -Command "gpupdate /force"
   }
   Invoke-Command -ComputerName $computer -ScriptBlock $scriptblock
}