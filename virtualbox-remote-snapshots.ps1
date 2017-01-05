# We get passed the pwd, so we can look for the config file
param(
  [string]$Pwd
)

# Self-elevating stub, since VSS snapshots can only be created by an administrator
If (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] 'Administrator')) {
  $Pwd = (Convert-Path .)
  Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`" $Pwd" -Verb RunAs; exit
}

Write-Host 'Probing for VMs...'

$VBoxManage = 'C:\Program Files\Oracle\VirtualBox\VBoxManage'

$VMs = (& $VBoxManage list vms)
For ($I = 0; $I -lt $VMs.count; $I++) {
  $VMs[$I] = [regex]::match($VMs[$I], '\"(.*?)\"').Groups[1].Value
  Write-Host ('[{0}] {1}' -f $I, $VMs[$I])
}

While ($true) {
  $VMIndex = ((Read-Host 'Press select a VM by number') -as [int])
  If (0 -le $VMIndex -le $VMs.count - 1) {
    $VMname = $VMs[$VMIndex]
    Break
  }
  Else {
    Write-Host 'Index out of range!'
  }
}

Write-Host 'Detecting VM location...'

$VMInfo = (& $VBoxManage showvminfo $VMName)

$VMConfigFile = ($VMInfo | Select-String 'Config file:')
$VMConfigFile = [regex]::match($VMConfigFile, 'Config file:\s+(.*)$')
$VMConfigFile = $VMConfigFile.Groups[1].Value

$VMLocation = ($VMConfigFile -replace '[^\\]*$')
$VSSDriveLetter = $VMConfigFile.Substring(0, 1)
$VSSTarget = '{0}:\.atomic' -f $VSSDriveLetter

$BorgTarget = '/cygdrive/{0}/.atomic' -f $VSSDriveLetter.ToLower()
$BorgTarget = $BorgTarget + ($VMLocation -replace '.:')
$BorgTarget = $BorgTarget -replace '\\', '/' -replace ' ', '\\ '
$BorgArchiveTag = $VMName -replace ' ','-'

Write-Host "VM located at: $BorgTarget"

# EOL hack to stop Cygwin from complaining about carriage returns
$EOL=";: "

# Parse the config file
Get-Content "$Pwd\virtualbox-remote-snapshots.conf" | ForEach-Object -begin {$Conf=@{}} -process {
  $Key = [regex]::split($_, '=')
  if(($Key[0] -ne '') -and ($Key[0].StartsWith("[") -ne $True)) {
    $Conf.Add($Key[0], $Key[1])
  }
}

Write-Host 'Config:'
Write-Host ($Conf | Out-String)

$BackupHost = $Conf.BackupHost
$BackupHostPath = $Conf.BackupHostPath

# Query the user for the backup password
$SecurePassword = Read-Host -Prompt 'Enter password' -AsSecureString

While ($true) {
  # Cleanup the VSS link target, just in case
  If (Test-Path "${VSSTarget}") {
    cmd /c "rmdir /s /q ${VSSTarget}"
  }

  # Create a new VSS snapshot and detect its ID
  $ShadowStatus = (wmic shadowcopy call create Volume="${VSSDriveLetter}:\")
  $ShadowID = [regex]::match($ShadowStatus, 'ShadowID = "([^"]+)').Groups[1].Value

  Write-Output "Sucessfully created shadow copy at ${ShadowID}."

  Try {
    # Capture the snapshot's Device Object and "mount" it to VSSTarget
    $DeviceObject = ([regex]::match((wmic shadowcopy get DeviceObject,ID), "(\S+)\s+$ShadowID").Groups[1].Value) + "\"
    cmd /c "mklink /d ${VSSTarget} ${DeviceObject}"

    # Check if VMs are online
    $VMsOffline = (& 'C:\Program Files\Oracle\VirtualBox\VBoxManage' list runningvms) -eq $null
    $StateSuffix = If ($VMsOnline) { 'offline' } Else { 'online' }

    # Decrypt the password in memory
    $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecurePassword)
    $PlainPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
    Remove-Variable BSTR

    # Start the backup
    & "C:\Program Files\Borg\bin\bash.exe" -l -c @"
export BORG_PASSPHRASE=${PlainPassword}${EOL}
echo Starting borg backup...${EOL}
cd $BorgTarget${EOL}
borg create -vspx -C lz4 ${BackupHost}:${BackupHostPath}::'$BorgArchiveTag-{now:%Y.%m.%d-%H.%M.%S}-$StateSuffix' .${EOL}
echo Starting borg prune...${EOL}
borg prune -vs --list ${BackupHost}:${BackupHostPath} --keep-within 2H -H 8 -d 7 -w 3${EOL}
"@j
    Remove-Variable PlainPassword
  }
  Catch {
    Write-Output 'An error has occured. Cleaning up...'
  }
  Finally {
    # Remove the VSS target and delete the VSS snapshot
    cmd /c "rmdir /s /q ${VSSTarget}"
    vssadmin delete shadows /shadow=$ShadowID /quiet
    Write-Output "Successfully deleted shadow copy ${ShadowID}."
  }

  Read-Host 'Press Enter to backup again...' | Out-Null
}
