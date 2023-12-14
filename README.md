# Backup-VirtualBox-PowerShell
A simple PowerShell script that allows you to backup VirtualBox machines. They are being exported using the VBoxManage system that VirtualBox uses. It (gently) stops the machine, takes an OVA and starts it back up. 

## Switches
* Compress, to allow compression of the backup
* StartAfterBackup, this allows you to start the VM back up after the backup
* Force, used to tell VirtualBox to force shutdown (poweroff) the VM instead of acpipowerbutton method

## Prereq
* 7-Zip is needed to do the compression if you should choose that option from the script this is done because the archiving functions of .NET do not allow archiving of large files. 

## Learning
The script is me learning PowerShell and doing day to day tasks. I am not a PowerShell pro but I hope this will help you. 
