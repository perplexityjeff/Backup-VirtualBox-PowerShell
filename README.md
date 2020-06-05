# Backup-VirtualBox-PowerShell
A simple PowerShell script that allows you to backup VirtualBox machines. They are being exported using the VBoxManage system that VirtualBox uses. It (gently) stops the machine, takes an OVA and starts it back up. 

# Prereq
* 7-Zip is needed to do the compression if you should choose that option from the script this is done because the archiving functions of .NET do not allow for to large of files. 
