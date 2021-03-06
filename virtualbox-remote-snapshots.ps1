# EOL hack to stop Cygwin from complaining about carriage returns
$EOL=";: "
$PersistFile="./.vrs_persist"

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
$BorgFlags = $Conf.BorgFlags
$RsyncFlags = $Conf.RsyncFlags

# Query the user for the backup password
Write-Host 'Enter password: ' -NoNewline -Fore magenta
$SecurePassword = Read-Host -AsSecureString

$VMName = ""
$VMLocation = ""
$VSSDriveLetter = ""
$VSSTarget = ""
$BorgTarget = ""
$BorgArchiveTag = ""

$VBoxManage = 'C:\Program Files\Oracle\VirtualBox\VBoxManage'
$BorgBash = 'C:\Program Files\Borg\bin\bash.exe'
$SSHArgs = '-Tx -c aes128-gcm@openssh.com'
$SSHBase = "ssh ${SSHArgs} ${BackupHost}"

Function Select-VM () {
  If (-not (Test-Path $PersistFile)) {
    New-Item $PersistFile -type file
  }
  Write-Host 'Probing for VMs...'

  $VMs = @(& $VBoxManage list vms)

  For ($I = 0; $I -lt $VMs.count; $I++) {
    $VMs[$I] = [regex]::match($VMs[$I], '\"(.*?)\"').Groups[1].Value
  }

  If ([string]::ISNullOrEmpty($Global:VMName)) {
    Try {
      $Global:VMName = $VMs[(Get-Content $PersistFile)]
    } Catch {
      $Global:VMName = $VMs[0]
    }
  }
  ElseIf ($VMs.Count -eq 2) {
    # just switch to the other VM
    If ($VMs[1] -eq $Global:VMName) {
      $Global:VMName = $VMs[0]
      Set-Content $PersistFile 0
    }
    Else {
      $Global:VMName = $VMs[1]
      Set-Content $PersistFile 1
    }
  }
  Else {
    For ($I = 0; $I -lt $VMs.count; $I++) {
      Write-Host ('[{0}] {1}' -f $I, $VMs[$I])
    }

    While ($true) {
      Write-Host 'Select a VM by number: ' -nonewline -fore magenta
      $VMIndex = Read-Host
      Try {
        $VMIndex = $VMIndex -as [int]
      }
      Catch {
        Continue
      }
      If (0 -le $VMIndex -le $VMs.count - 1) {
        Set-Content $PersistFile $VMIndex
        $Global:VMname = $VMs[$VMIndex]
        Break
      }
      Else {
        Write-Host 'Index out of range!'
      }
    }
  }

  Write-Host 'Detecting VM location...'

  $VMInfo = @(& $VBoxManage showvminfo $Global:VMName --machinereadable)

  $VMConfigFile = ($VMInfo | Select-String 'CfgFile=')
  $VMConfigFile = [regex]::match($VMConfigFile, 'CfgFile="(.*)"$')
  $VMConfigFile = $VMConfigFile.Groups[1].Value

  $Global:VMLocation = ($VMConfigFile -replace '[^\\]*$' -replace '\\\\','\')
  $Global:VSSDriveLetter = $VMConfigFile.Substring(0, 1)
  $Global:VSSTarget = '{0}:\.atomic' -f $Global:VSSDriveLetter

  $Global:BorgTarget = '/cygdrive/{0}/' -f $Global:VSSDriveLetter.ToLower()
  $Global:BorgTarget = $Global:BorgTarget + ($Global:VMLocation -replace '.:')
  $Global:BorgTarget = $Global:BorgTarget -replace '\\', '/' -replace ' ', '\\ '
  $Global:BorgArchiveTag = $Global:VMName -replace ' ','-'

  Write-Host "VM located at: $Global:VMLocation" -fore cyan
}

Function Snapshot-Create () {
  # Check if VMs are online
  $VMOnline = @(& $VBoxManage list runningvms | Select-String $Global:VMName).Count -ne 0

  If ($VMOnline) {
    Write-Host 'Taking online snapshot...' -fore cyan
    & $VBoxManage snapshot $Global:VMname take "vrs-temp" --live
  }
  Else {
    & $VBoxManage snapshot $Global:VMname take "vrs-temp"
  }

  Try {
    $StateSuffix = If ($VMOnline) { 'online' } Else { 'offline' }
    # Decrypt the password in memory
    $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecurePassword)
    $PlainPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
    Remove-Variable BSTR

    # Start the backup
    & $BorgBash -l -c @"
export BORG_PASSPHRASE=${PlainPassword}${EOL}
export BORG_RSH='ssh ${SSHArgs}'${EOL}
echo Starting borg backup...${EOL}
cd ${Global:BorgTarget}${EOL}
borg create -vspx -C lz4 ${BorgFlags} ${BackupHost}:${BackupHostPath}::'${Global:BorgArchiveTag}-{now:%Y.%m.%d-%H.%M.%S}-$StateSuffix' .${EOL}
"@
    Remove-Variable PlainPassword
  }
  Catch {
    Write-Output $_.Exception.Message
    Write-Output 'An error has occured. Cleaning up...'
  }
  Finally {
    Write-Host 'Deleting VirtualBox Snapshot...' -fore cyan
    & $VBoxManage snapshot $Global:VMname delete "vrs-temp"
  }
}

Function Snapshot-Migrate () {
  # Check if VMs are online
  $VMOnline = @(& $VBoxManage list runningvms | Select-String $Global:VMName).Count -ne 0

  If ($VMOnline) {
    Write-Host 'Saving Vm state...' -fore cyan
    & $VBoxManage controlvm $Global:VMname savestate
  }

  Try {
    $StateSuffix = If ($VMOnline) { 'online' } Else { 'offline' }
    # Decrypt the password in memory
    $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecurePassword)
    $PlainPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
    Remove-Variable BSTR

    # Start the backup
    & $BorgBash -l -c @"
export BORG_PASSPHRASE=${PlainPassword}${EOL}
export BORG_RSH='ssh ${SSHArgs}'${EOL}
echo Starting borg backup...${EOL}
cd ${Global:BorgTarget}${EOL}
borg create -vspx -C lz4 ${BorgFlags} ${BackupHost}:${BackupHostPath}::'${Global:BorgArchiveTag}-{now:%Y.%m.%d-%H.%M.%S}-$StateSuffix' .${EOL}
"@
    Remove-Variable PlainPassword
  }
  Catch {
    Write-Output $_.Exception.Message
    Write-Output 'An error has occured. Cleaning up...'
  }
}

Function Snapshot-Prune () {
  Try {
    # Decrypt the password in memory
    $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecurePassword)
    $PlainPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
    Remove-Variable BSTR

    # Start the backup
    & $BorgBash -l -c @"
export BORG_PASSPHRASE=${PlainPassword}${EOL}
export BORG_RSH='ssh ${SSHArgs}'${EOL}
echo Starting borg prune...${EOL}
borg prune -vs --list ${BackupHost}:${BackupHostPath} -P '${Global:BorgArchiveTag}' --keep-within 2H -H 8 -d 7 -w 3${EOL}
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

  $ExtractTarget = ($Global:BorgTarget -replace 'cygdrive', 'mnt')

  # Decrypt the password in memory
  $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecurePassword)
  $PlainPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
  Remove-Variable BSTR

  # Mount the borg repository
  Write-Host "Getting list of Archives..."
  $MountpointTest=@(& $BorgBash -l -c "${SSHBase} 'cd ${RemoteMountpoint}; find . -maxdepth 1 ! -path . -printf \`"%f\n\`"'")
  If ($MountpointTest.Count -ne 0) {
    & $BorgBash -l -c "${SSHBase} 'fusermount -u ${RemoteMountpoint}'"
  }

  Write-Host "Verified mountpoint"

  Try {
    & $BorgBash -l -c "${SSHBase} 'BORG_PASSPHRASE=${PlainPassword} borg mount -o allow_other ${BackupHostPath} ${RemoteMountpoint}'"
    Remove-Variable PlainPassword

    Write-Host "Mount completed"
    $Archives=@(& $BorgBash -l -c "${SSHBase} 'find ${RemoteMountpoint} -name \`"${Global:BorgArchiveTag}*\`" -maxdepth 1 ! -path . -printf \`"%f\n\`" | sort -n'")

    Write-Host "Archive list retrieved"
    Write-Host
    Write-Host "Select archive to restore:"

    For ($I = 0; $I -lt $Archives.count; $I++) {
      Write-Host ("[{0,2}] {1}" -f $I, $Archives[$I])
    }

    # TODO: Validate this input
    $ArchiveLatest = $Archives.count - 1
    While ($true) {
      Write-Host "Enter archive number (default is `"${ArchiveLatest}`"): " -nonewline -fore magenta
      $ArchiveSelected = Read-Host
      If ($ArchiveSelected -eq "") {
        $ArchiveSelected = $ArchiveLatest
        Break
      }
      Else {
        Try {
          $ArchiveSelected = $ArchiveSelected -as [int]
        }
        Catch {
          Return
        }
        If (0 -le $ArchiveSelected -le $Archives.count - 1) {
          Break
        }
      }
    }
    $Archive = $Archives[$ArchiveSelected]
    Write-Host ("Restoring archive {0}..." -f $Archive)

    $Path = ""

    While ($true) {
      $DirScan = @(& $BorgBash -l -c "${SSHBase} 'find ${RemoteMountpoint}/${Archive}/${Path} -maxdepth 1 ! -path . -printf \`"%f\n\`"'")

      If (($DirScan -like '*.vbox').count -eq 1) {
        Break
      }
      ElseIf ($DirScan.Count -gt 1) {
        Throw "Could not detect VM config file in archive!"
      }

      $Path = $Path + ($DirScan[0] -replace " ","\ ") + "/"
    }

    Write-Host ("Detected VM folder: {0}." -f $Path)

    $VMOnline = @(& $VBoxManage list runningvms | Select-String $Global:VMName).Count -ne 0
    If ($VMOnline) {
      & $VBoxManage controlvm $Global:VMname poweroff
    }

    & $VBoxManage unregistervm $Global:VMname

    Write-Host "Unregistered VM" -fore cyan

    bash -c "time rsync -vah -e 'ssh ${SSHArgs}' --inplace --info=progress2 --stats --delete-before ${RsyncFlags} ${BackupHost}:'${RemoteMountpoint}/${Archive}/${Path}' ${ExtractTarget}/"
  }
  Catch {
    Write-Output $_.Exception.Message
    Write-Output "An error has occured. Cleaning up..."
  }
  Finally {
    $VBoxFiles = @(Get-ChildItem ${Global:VMLocation}/*.vbox)
    If ($VBoxFiles.Count -ne 1) {
      Throw "Could not detect VM config file in extracted archive!"
    }

    $VBoxFileName = $VBoxFiles[0].Name
    & $VBoxManage registervm "${Global:VMLocation}/${VBoxFileName}"
    Write-Host "Registered VM" -fore cyan

    & $BorgBash -l -c "${SSHBase} 'fusermount -u ${RemoteMountpoint}'"
    Write-Host "Unmounted filesystem"
  }

  $VMSnapshots = @(& $VBoxManage snapshot $Global:VMName list | Select-String 'vrs-temp')
  If ($VMSnapshots.Count -ne 0) {
    Write-Host 'Resolving snapshots...'
    & $VBoxManage snapshot $Global:VMname restore 'vrs-temp'
    & $VBoxManage snapshot $Global:VMName delete 'vrs-temp'
  }

}

Select-VM

While ($true) {
  $ActionTitle = "Select Action [$Global:VMName]"
  $ActionMessage = 'What action would you like to take?'

  $ActionSnapshot = New-Object System.Management.Automation.Host.ChoiceDescription `
    '&Snapshot', 'Take a snapshot of this virtual machine.'

  $ActionMigrate = New-Object System.Management.Automation.Host.ChoiceDescription `
    '&Migrate', 'Migrate this virtual machine to another host.'

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
    $ActionMigrate, `
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
      Snapshot-Migrate
    }
    2 {
      Snapshot-Prune
    }
    3 {
      Snapshot-Restore
    }
    4 {
      Select-VM
    }
    5 {
      Exit
    }
  }
}
