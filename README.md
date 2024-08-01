# Backup-VirtualBox-PowerShell
A simple PowerShell script that allows you to backup VirtualBox machines. 

They are being exported using the VBoxManage command line that comes with VirtualBox. It (gently) stops the machine, takes an OVA and starts it back up. 

## Usage
Here is a simple example of the script that you use to backup a VM:

`.\Backup-VirtualBox.ps1 -VM 'VM Test' -Destination D:\Backup -Compress -StartAfterBackup -Verbose`

Please use Get-Help for more examples within the script:

`Get-Help .\Backup-VirtualBox.ps1 -Examples`

## Mode: FullBackup
This mode is shutting down the VM (gracefully) and exporting the VM to a OVA. 

* VM: This is the VM name as displayed in VirtualBox itself
* Destination: This is the folder in which the backup will be stored
* Suffix: This gets added after the filename generated
* Compress: This allow you to have compression of the backup using [7-Zip](https://www.7-zip.org/)
* CompressExtension: This allows you to change the compression type e.g "zip" or "7z" and is optional
* CompressLevel: This allows you to change the compression level e.g "1", "3", "5" (Default), "7", "9"
* StartAfterBackup: This allows you to start the VM back up after the backup
* Force: This is used to tell VirtualBox to force shutdown (poweroff) the VM instead of acpipowerbutton method

## Mode: Snapshot
This only creates a snapshot of the VM

* VM: This is the VM name as displayed in VirtualBox itself
* Snapshot: This is a switch to get into snapshot mode. This does not save it to a destination thus not a backup
* Keep: This allows you to specify the amount of days of snapshots to keep, older snapshots will be deleted. This function is optional

## Filenames
The filename of the backup or snapshots by default is the VM name and a date timestamp (yyyyMMdd-HHmmss). Using the optional Suffix parameter for backups you can add something after that, for example "Daily". 

## Prerequisite
* [7-Zip](https://www.7-zip.org/) is needed to do the compression if you should choose that option from the script this is done because the archiving functions of .NET do not allow archiving of large files. 

## Learning
The script is me learning PowerShell and doing day to day tasks or just for the fun of it. I am not a PowerShell pro but I hope this will help someone. 
