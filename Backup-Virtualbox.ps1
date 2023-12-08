<#
.SYNOPSIS
This is a simple Powershell script to allow you to create backups of VirtualBox VM's using the 'Export to Appliance' command. 

.DESCRIPTION
The script is using the command line interface of VirtualBox named VBoxManage to do the starting and stopping of the VM's and the creation of the exported OVA files. 
It is as well required to have 7-Zip installed on the computer. 

It uses 7-Zip to compress the backupped directory. The command line argument Compress-Archive is not 
supported as the output files are most likely larger than 2GB. Because of a .NET limitation zip archives larger then 2GB are not supported yet because of Zip64 not being 
implemented in .NET yet. 

Make sure to use StartAfterBackup if you plan to stop the VM, make a backup and start the VM. This is so that you are able to make backups of machines that are offline
and should stay offline.

.EXAMPLE
.\Backup-VirtualBox.ps1 -VM 'TESTVM' -Destination D:\Test\TESTVM -Verbose

.EXAMPLE
.\Backup-VirtualBox.ps1 -VM 'TESTVM' -Destination D:\Test\TESTVM -Compress -Verbose

.EXAMPLE
.\Backup-VirtualBox.ps1 -VM 'TESTVM' -Destination D:\Test\TESTVM -Compress -StartAfterBackup -Verbose

.LINK
https://www.techrepublic.com/article/how-to-import-and-export-virtualbox-appliances-from-the-command-line/
https://www.virtualbox.org/manual/ch08.html
https://www.virtualbox.org/manual/ch08.html#vboxmanage-export
https://stackoverflow.com/questions/49807310/in-powershell-with-a-large-number-of-files-can-you-send-2gb-at-a-time-to-zip-fil
https://blogs.msdn.microsoft.com/oldnewthing/20180515-00/?p=98755
https://stackoverflow.com/questions/17461237/how-do-i-get-the-directory-of-the-powershell-script-i-execute
#>

[cmdletBinding()]
Param
(
    [Parameter(Mandatory=$true)][String]$VM = "",
    [Parameter(Mandatory=$true)][String]$Destination = "C:\Users\" + $env:UserName + "\Documents\",
    [Switch]$Compress,
    [Switch]$StartAfterBackup
)

function Create-7Zip([String] $aDirectory, [String] $aZipfile){
    [string]$pathToZipExe = "$($Env:ProgramFiles)\7-Zip\7z.exe";
    [Array]$arguments = "a", "-tzip", "$aZipfile", "$aDirectory", "-r";
    & $pathToZipExe $arguments;
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

$Date = Get-Date -format "yyyyMMdd"
$VBoxManage = 'C:\Program Files\Oracle\VirtualBox\VBoxManage.exe'
$OVA = "$VM-$Date.ova"
$OVAPath = $PSScriptRoot + "\" + $OVA

Write-Verbose "Stopping $VM"
Start-Process $VBoxManage -ArgumentList "controlvm ""$VM"" poweroff" -Wait -WindowStyle Hidden

Write-Verbose "Testing if $Destination exists, if not then create it"
if (-Not(Test-Path $Destination))
{
    New-Item -Path $Destination -ItemType Directory
}

Write-Verbose "Checking if $OVA already exists and removing it before beginning"
if (Test-Path $OVAPath)
{
    Remove-Item $OVAPath -Force -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent)
}

Write-Verbose "Waiting for $VM to have stopped"
While(Get-RunningVirtualBox($VM))
{
    Start-Sleep -Seconds 1
}

Write-Verbose "Exporting the VM appliance of $VM as $OVA"
Start-Process $VBoxManage -ArgumentList "export ""$VM"" -o ""$OVAPath""" -Wait -WindowStyle Hidden

if ($StartAfterBackup)
{
    Write-Verbose "Starting $VM"
    Start-Process $VBoxManage -ArgumentList "startvm ""$VM"" -type headless" -Wait -WindowStyle Hidden
}

if ($Compress)
{
    $DestinationCompress = $Destination + "\" + $OVA.Split('.')[0] + ".zip"
   
    Write-Verbose "Checking if $DestinationCompress already exists and removing it before beginning"
    if (Test-Path $DestinationCompress)
    {
        Remove-Item $DestinationCompress -Force -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent)
    }

    Write-Verbose "Starting the compression of $OVA to $DestinationCompress"
    Create-7Zip ($OVAPath) $DestinationCompress

    Write-Verbose "Removing $OVAPath because of completed compression"
    Remove-Item ($OVAPath) -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent)
}
else
{
    Write-Verbose "Copying the exported $OVA to $Destination"
    Copy-Item ($OVAPath) -Destination "($Destination + "\" + $OVA)" -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent)
}

Write-Verbose "Completed the Backup"
