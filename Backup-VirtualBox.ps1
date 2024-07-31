<#
.SYNOPSIS
This is a simple Powershell script to allow you to create backups of VirtualBox VM's using the 'Export to Appliance' command. It uses the VM name as displayed in VirtualBox Manager as the input 

.DESCRIPTION
The script is using the command line interface of VirtualBox named VBoxManage to do the starting and stopping of the VM's and the creation of the exported OVA files. 
It is as well required to have 7-Zip installed on the computer. 

It uses 7-Zip to compress the backupped directory. The command line argument Compress-Archive is not 
supported as the output files are most likely larger than 2GB. Because of a .NET limitation zip archives larger then 2GB are not supported yet because of Zip64 not being 
implemented in .NET yet. 

Make sure to use StartAfterBackup if you plan to stop the VM, make a backup and start the VM. This is so that you are able to make backups of machines that are offline
and should stay offline.

.EXAMPLE
Basic example to backup the VM to a folder
.\Backup-VirtualBox.ps1 -VM 'TESTVM' -Destination D:\Test\TESTVM -Verbose

.EXAMPLE
Basic example to backup the VM to a folder and use zip archiving with 7zip
.\Backup-VirtualBox.ps1 -VM 'TESTVM' -Destination D:\Test\TESTVM -Compress -Verbose

.EXAMPLE
Basic example to backup the VM to a folder and use 7z archiving with 7zip
.\Backup-VirtualBox.ps1 -VM 'TESTVM' -Destination D:\Test\TESTVM -Compress -CompressExtention "7z" -Verbose

.EXAMPLE
Basic example to backup the VM to a folder and use 7z archiving with 7zip with maximum compression 9 (this can also be 1, 3, 5 (default) or 7)
.\Backup-VirtualBox.ps1 -VM 'TESTVM' -Destination D:\Test\TESTVM -Compress -CompressExtention "7z" -CompressLevel 9 -Verbose

.EXAMPLE
Basic example to backup the VM to a folder and start the VM back up after exporting
.\Backup-VirtualBox.ps1 -VM 'TESTVM' -Destination D:\Test\TESTVM -Compress -StartAfterBackup -Verbose

.LINK
https://www.techrepublic.com/article/how-to-import-and-export-virtualbox-appliances-from-the-command-line/
https://www.virtualbox.org/manual/ch08.html
https://www.virtualbox.org/manual/ch08.html#vboxmanage-export
https://stackoverflow.com/questions/49807310/in-powershell-with-a-large-number-of-files-can-you-send-2gb-at-a-time-to-zip-fil
https://blogs.msdn.microsoft.com/oldnewthing/20180515-00/?p=98755
https://stackoverflow.com/questions/17461237/how-do-i-get-the-directory-of-the-powershell-script-i-execute
https://askubuntu.com/questions/42482/how-to-safely-shutdown-guest-os-in-virtualbox-using-command-line
https://www.dotnetperls.com/7-zip-examples
#>

[CmdletBinding()]
Param
(
    [Parameter(Mandatory=$true)][String]$VM = "",
    [Parameter(Mandatory=$true)][String]$Destination = "C:\Users\" + $env:UserName + "\Documents\",
    [String]$Suffix,
    [Switch]$Compress,
    [String]$CompressExtension = "zip",
    [String]$CompressLevel = "5",    
    [Switch]$StartAfterBackup,
    [switch]$Force = $False
)

function New-7ZipArchive()
{
    Param
    (
        [Parameter(Mandatory=$true)][String]$DestinationFile, 
        [Parameter(Mandatory=$true)][String]$SourceFile, 
        [Parameter(Mandatory=$true)][String]$Extention, 
        [Parameter(Mandatory=$true)][String]$Level
    )

    $7ZipLocation = "$($Env:ProgramFiles)\7-Zip\7z.exe"
    $7ZipArguments = "a -t" + $Extention + ' "' + $DestinationFile + '" "' + $SourceFile + '" -mx' + $Level

    Start-Process $7ZipLocation -ArgumentList $7ZipArguments -Wait -WindowStyle Hidden
}

function Get-RunningVirtualBox($VM)
{
    $VBoxManage = 'C:\Program Files\Oracle\VirtualBox\VBoxManage.exe'

    $pinfo = New-Object System.Diagnostics.ProcessStartInfo
    $pinfo.FileName = $VBoxManage
    $pinfo.RedirectStandardError = $true
    $pinfo.RedirectStandardOutput = $true
    $pinfo.UseShellExecute = $false
    $pinfo.Arguments = "list runningvms"
    $p = New-Object System.Diagnostics.Process
    $p.StartInfo = $pinfo
    $p.Start() | Out-Null
    $p.WaitForExit()
    $stdout = $p.StandardOutput.ReadToEnd()
    $stderr = $p.StandardError.ReadToEnd()

    if ($stdout.Contains($VM))
    {
        return $stdout
    }
}

$Date = Get-Date -format "yyyyMMdd-HHmmss"
$VBoxManage = 'C:\Program Files\Oracle\VirtualBox\VBoxManage.exe'

$OVA = "$VM-$Date.ova"
if ($Suffix)
{
    $OVA = "$VM-$Date-$Suffix.ova"
}

$OVAPath = Join-Path -Path $Destination -ChildPath $OVA

if (Get-RunningVirtualBox($VM))
{
    if ($Force)
    {
        Write-Verbose "Stopping $VM, using PowerOff method (Force)"
        Start-Process $VBoxManage -ArgumentList "controlvm ""$VM"" poweroff" -Wait -WindowStyle Hidden
    }
    else
    {
        Write-Verbose "Stopping $VM, using ACPI Power Button method"
        Start-Process $VBoxManage -ArgumentList "controlvm ""$VM"" acpipowerbutton" -Wait -WindowStyle Hidden
    }
   
    While(Get-RunningVirtualBox($VM))
    {
        Write-Verbose "Waiting for $VM to have stopped"
        Start-Sleep -Seconds 1
    }
}

Write-Verbose "Testing if $Destination exists, if not then create it"
if (-Not(Test-Path $Destination))
{
    New-Item -Path $Destination -ItemType Directory | Out-Null
}

Write-Verbose "Checking if $OVA already exists and removing it before beginning"
if (Test-Path $OVAPath)
{
    Remove-Item $OVAPath -Force
}

Write-Verbose "Exporting the VM appliance of $VM as $OVA, be patient this will take a while"
Start-Process $VBoxManage -ArgumentList "export ""$VM"" -o ""$OVAPath""" -Wait -WindowStyle Hidden -RedirectStandardOutput True

if ($StartAfterBackup)
{
    Write-Verbose "Starting $VM"
    Start-Process $VBoxManage -ArgumentList "startvm ""$VM"" -type headless" -Wait -WindowStyle Hidden
}

if ($Compress)
{
    $DestinationCompress = Join-Path $Destination -ChildPath ((Get-Item $OVAPath).BaseName + "." + $CompressExtension)
   
    Write-Verbose "Checking if $DestinationCompress already exists and removing it before beginning"
    if (Test-Path $DestinationCompress)
    {
        Remove-Item $DestinationCompress -Force
    }

    Write-Verbose "Starting the compression of $OVA to $DestinationCompress"
    New-7ZipArchive -SourceFile $OVAPath -DestinationFile $DestinationCompress -Extention $CompressExtension -Level $CompressLevel

    Write-Verbose "Removing uncompressed $OVAPath because of completed compression"
    Remove-Item $OVAPath -Force
}

Write-Verbose "Completed the backup"
