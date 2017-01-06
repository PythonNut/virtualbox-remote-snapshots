# We get passed the pwd, so we can look for the config file
param(
  [string]$Pwd
)

# Self-elevating stub, since VSS snapshots can only be created by an administrator
If (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] 'Administrator')) {
  $Pwd = (Convert-Path .)
  Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`" $Pwd" -Verb RunAs; exit
}

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
$RemoteMountpoint = $Conf.RemoteMountpoint

# Query the user for the backup password
$SecurePassword = Read-Host -Prompt 'Enter password' -AsSecureString

$VMName = ""
$VMLocation = ""
$VSSDriveLetter = ""
$VSSTarget = ""
$BorgTarget = ""
$BorgArchiveTag = ""

Function Select-VM () {
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
      $Global:VMname = $VMs[$VMIndex]
      Break
    }
    Else {
      Write-Host 'Index out of range!'
    }
  }

  Write-Host 'Detecting VM location...'

  $VMInfo = (& $VBoxManage showvminfo $Global:VMName)

  $VMConfigFile = ($VMInfo | Select-String 'Config file:')
  $VMConfigFile = [regex]::match($VMConfigFile, 'Config file:\s+(.*)$')
  $VMConfigFile = $VMConfigFile.Groups[1].Value

  $Global:VMLocation = ($VMConfigFile -replace '[^\\]*$')
  $Global:VSSDriveLetter = $VMConfigFile.Substring(0, 1)
  $Global:VSSTarget = '{0}:\.atomic' -f $Global:VSSDriveLetter

  $Global:BorgTarget = '/cygdrive/{0}/.atomic' -f $Global:VSSDriveLetter.ToLower()
  $Global:BorgTarget = $Global:BorgTarget + ($Global:VMLocation -replace '.:')
  $Global:BorgTarget = $Global:BorgTarget -replace '\\', '/' -replace ' ', '\\ '
  $Global:BorgArchiveTag = $Global:VMName -replace ' ','-'

  Write-Host "VM located at: $BorgTarget"
}

Function Snapshot-Create () {
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
"@
    Remove-Variable PlainPassword
  }
  Catch {
    Write-Output $_.Exception.Message
    Write-Output 'An error has occured. Cleaning up...'
  }
  Finally {
    # Remove the VSS target and delete the VSS snapshot
    cmd /c "rmdir /s /q ${VSSTarget}"
    vssadmin delete shadows /shadow=$ShadowID /quiet
    Write-Output "Successfully deleted shadow copy ${ShadowID}."
  }
}

Function Snapshot-Prune () {
  Try {
    # Decrypt the password in memory
    $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecurePassword)
    $PlainPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
    Remove-Variable BSTR

    # Start the backup
    & "C:\Program Files\Borg\bin\bash.exe" -l -c @"
export BORG_PASSPHRASE=${PlainPassword}${EOL}
echo Starting borg prune...${EOL}
borg prune -vs --list ${BackupHost}:${BackupHostPath} -P '${BorgArchiveTag}' --keep-within 2H -H 8 -d 7 -w 3${EOL}
"@
    Remove-Variable PlainPassword
  }
  Catch {
    Write-Output $_.Exception.Message
    Write-Output 'An error has occured. Cleaning up...'
  }
}

Function Snapshot-Restore () {
  # TODO: Ensure VM is powered off before restoring

  $ExtractTarget = ($BorgTarget -replace '\.atomic/' -replace 'cygdrive','mnt')

  # Decrypt the password in memory
  $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecurePassword)
  $PlainPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
  Remove-Variable BSTR

  # Mount the borg repository
  Write-Host "Getting list of Archives..."
  $MountpointTest=(& "C:\Program Files\Borg\bin\bash.exe" -l -c "ssh ${BackupHost} 'ls -1 ${RemoteMountpoint}'")
  If ($MountpointTest.count -gt 0) {
    & "C:\Program Files\Borg\bin\bash.exe" -l -c "ssh ${BackupHost} 'fusermount -u ${RemoteMountpoint}'"
  }

  Write-Host "Verified mountpoint."

  Try {
    & "C:\Program Files\Borg\bin\bash.exe" -l -c "ssh ${BackupHost} 'BORG_PASSPHRASE=${PlainPassword} borg mount -o allow_other ${BackupHostPath} ${RemoteMountpoint}'"
    Remove-Variable PlainPassword

    Write-Host "Mount completed"
    $Archives=@(& "C:\Program Files\Borg\bin\bash.exe" -l -c "ssh ${BackupHost} 'ls -1d ${RemoteMountpoint}/${BorgArchiveTag}*'")
    Write-Host "Archive list retrieved"
    Write-Host ""
    Write-Host "Select archive to restore:"
    For ($I = 0; $I -lt $Archives.count; $I++) {
      Write-Host ("[{0:00}] {1}" -f $I, $Archives[$I])
    }

    # TODO: Validate this input
    $Archive = $Archives[(Read-Host 'Enter archive number') -as [int]]
    Write-Host ("Restoring archive {0}..." -f $Archive)

    $Path = ""

    While ($true) {
      $DirScan = @(& "C:\Program Files\Borg\bin\bash.exe" -l -c "ssh ${BackupHost} 'ls -A1 ${Archive}/${Path}'")

      If (($DirScan -like '*.vbox').count -eq 1) {
        Break
      }
      ElseIf ($DirScan.count -gt 1) {
        Throw "Could not detect VM config file in archive!"
      }

      $Path = $Path + ($DirScan[0] -replace " ","\ ") + "/"
    }

    Write-Host ("Detected VM folder: {0}." -f $Path)

    bash -c "time rsync -vah -e 'ssh -Tx -c aes128-gcm@openssh.com -o Compression=no' --inplace --info=progress2 --stats --delete-before ${BackupHost}:'${Archive}/${Path}' ${ExtractTarget}/"
  }
  Catch {
    Write-Output $_.Exception.Message
    Write-Output "An error has occured. Cleaning up..."
  }
  Finally {
    & "C:\Program Files\Borg\bin\bash.exe" -l -c "ssh ${BackupHost} 'fusermount -u ${RemoteMountpoint}'"
    Write-Host "Unmounted filesystem."
  }
}

Select-VM

While ($true) {
  $ActionTitle = "Select Action [$VMName]"
  $ActionMessage = 'What action would you like to take?'

  $ActionSnapshot = New-Object System.Management.Automation.Host.ChoiceDescription `
    '&Snapshot', 'Take a snapshot of this virtual machine.'

  $ActionPrune = New-Object System.Management.Automation.Host.ChoiceDescription `
    '&Prune', 'Prune old snapshots for this virtual machine.'

  $ActionRestore = New-Object System.Management.Automation.Host.ChoiceDescription `
    '&Restore', 'Restore a snapshot of this virtual machine.'

  $ActionSelect = New-Object System.Management.Automation.Host.ChoiceDescription `
    'S&elect', 'Select a different virtual machine.'

  $ActionQuit = New-Object System.Management.Automation.Host.ChoiceDescription `
    '&Quit', 'Exit from this session.'

  $ActionOptions = [System.Management.Automation.Host.ChoiceDescription[]]( `
    $ActionSnapshot, `
    $ActionPrune, `
    $ActionRestore, `
    $ActionSelect, `
    $ActionQuit)

  $ActionNumber = $host.ui.PromptForChoice( `
    $ActionTitle, `
    $ActionMessage, `
    $ActionOptions, `
    0)

  Switch ($ActionNumber) {
    0 {
      Snapshot-Create
    }
    1 {
      Snapshot-Prune
    }
    2 {
      Snapshot-Restore
    }
    3 {
      Select-VM
    }
    4 {
      Exit
    }
  }
}
