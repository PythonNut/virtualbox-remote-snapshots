# We get passed the pwd, so we can look for the config file
param(
  [string]$Pwd
)

Function Send-Tcp-Message ($Message=$([char]4), $Port=8998, $Server="localhost") {
  $Client = New-Object System.Net.Sockets.TcpClient $Server, $Port
  $Stream = $Client.GetStream()
  $Writer = New-Object System.IO.StreamWriter $Stream
  $Writer.Write($Message)
  $Writer.Dispose()
  $Stream.Dispose()
  $Client.Dispose()
}

# Self-elevating stub, since VSS snapshots can only be created by an administrator
If (([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] 'Administrator')) {

  $VSSDriveLetter = ""
  $VSSTarget = ""
  $ShadowID = ""

  Function Vss-Snapshot-Create () {
    # Cleanup the VSS link target, just in case
    If (Test-Path "${VSSTarget}") {
      cmd /c "rmdir /s /q ${VSSTarget}"
    }

    # Create a new VSS snapshot and detect its ID
    $ShadowStatus = (wmic shadowcopy call create Volume="${VSSDriveLetter}:\")
    $Global:ShadowID = [regex]::match($ShadowStatus, 'ShadowID = "([^"]+)').Groups[1].Value
    Try {
      Write-Output "Sucessfully created shadow copy at ${ShadowID}."

      # Capture the snapshot's Device Object and "mount" it to VSSTarget
      $DeviceObject = ([regex]::match((wmic shadowcopy get DeviceObject,ID), "(\S+)\s+$ShadowID").Groups[1].Value) + "\"
      cmd /c "mklink /d ${VSSTarget} ${DeviceObject}"
    }
    Catch { Vss-Snapshot-Delete }
  }

  Function Vss-Snapshot-Delete () {
    # Remove the VSS target and delete the VSS snapshot
    cmd /c "rmdir /s /q ${Global:VSSTarget}"
    vssadmin delete shadows /shadow=$Global:ShadowID /quiet
    Write-Output "Successfully deleted shadow copy ${Global:ShadowID}."
  }

  Function Listen-Port ($Port=8998) {
    $Endpoint = New-Object System.Net.IPEndPoint ([system.net.ipaddress]::any, $Port)
    $Listener = New-Object System.Net.Sockets.TcpListener $Endpoint
    $Listener.Start()

    Do {
      $Client = $Listener.AcceptTcpClient() # will block here until connection
      $Stream = $Client.GetStream();
      $Reader = New-Object System.IO.StreamReader $Stream
      Try {
        Do {
          $Line = $Reader.ReadLine()
          Write-Host $Line -fore cyan
          If ($Line -match 'VSSTarget=.*?') {
            $Global:VSSTarget = $Line -replace 'VSSTarget='
            Write-Host 'Set VSSTarget'
          }
          ElseIf ($Line -match 'VSSDriveLetter=.*?') {
            $Global:VSSDriveLetter = $Line -replace 'VSSDriveLetter='
            Write-Host 'Set VSSDriveLetter'
          }
          ElseIf ($Line -eq 'vss_create') {
            Vss-Snapshot-Create
            Send-Tcp-Message 'okay' 8231
          }
          ElseIf ($Line -eq 'vss_delete') {
            Vss-Snapshot-Delete
            Send-Tcp-Message 'okay' 8231
          }
          ElseIf ($Line -eq 'exit') {
            Exit
          }
        } While ($Line -and $Line -ne ([char]4))
      }
      Catch { }
      Finally {
        $Reader.Dispose()
        $Stream.Dispose()
        $Client.Dispose()
      }
    } While ($line -ne ([char]4))
    $Listener.Stop()
  }

  Try {
    Write-Host "VSS helper ready."

    Listen-Port 8230
  }
  Catch {
    Write-Output $_.Exception.Message
    Write-Host "An error occurred..."
  }
  Finally { Vss-Snapshot-Delete }
}
Else {
  Try {
    function Await-Tcp-Response ($Port=8998) {
      $Endpoint = New-Object System.Net.IPEndPoint ([system.net.ipaddress]::any, $Port)
      $Listener = New-Object System.Net.Sockets.TcpListener $Endpoint
      $Listener.Start()

      $Client = $Listener.AcceptTcpClient() # will block here until connection
      $Stream = $Client.GetStream();
      $Reader = New-Object System.IO.StreamReader $Stream
      Try {
        do {
          $Line = $Reader.ReadLine()
          if ($Line -match "okay") {
            break
          }
        } while ($Line -and $Line -ne ([char]4))
      }
      Catch { }
      Finally {
        $Reader.Dispose()
        $Stream.Dispose()
        $Client.Dispose()
        $Listener.stop()
      }
    }

    # TODO: I'm pretty sure we don't need this pwd nonsense anymore
    $Pwd = (Convert-Path .)
    Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`" $Pwd" -Verb RunAs
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

    Function Select-VM () {
      Write-Host 'Probing for VMs...'

      $VMs = (& $VBoxManage list vms)
      For ($I = 0; $I -lt $VMs.count; $I++) {
        $VMs[$I] = [regex]::match($VMs[$I], '\"(.*?)\"').Groups[1].Value
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
          $Global:VMname = $VMs[$VMIndex]
          Break
        }
        Else {
          Write-Host 'Index out of range!'
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

      Send-Tcp-Message "VSSTarget=${Global:VSSTarget}" 8230
      Send-Tcp-Message "VSSDriveLetter=${Global:VSSDriveLetter}" 8230

      $Global:BorgTarget = '/cygdrive/{0}/.atomic' -f $Global:VSSDriveLetter.ToLower()
      $Global:BorgTarget = $Global:BorgTarget + ($Global:VMLocation -replace '.:')
      $Global:BorgTarget = $Global:BorgTarget -replace '\\', '/' -replace ' ', '\\ '
      $Global:BorgArchiveTag = $Global:VMName -replace ' ','-'

      Write-Host "VM located at: $Global:BorgTarget"
    }

    Function Snapshot-Create () {
      # Check if VMs are online
      $VMOnline = @(& $VBoxManage list runningvms | Select-String $Global:VMName).Count -ne 0
      $StateSuffix = If ($VMOnline) { 'online' } Else { 'offline' }

      If ($VMOnline) {
        Write-Host 'Taking online snapshot...' -fore cyan
        & $VBoxManage snapshot $Global:VMname take "vrs-temp" --live
      }

      Try {
        Write-Host 'Creating VSS shadow...' -fore cyan
        Send-Tcp-Message 'vss_create' 8230
        Await-Tcp-Response 8231
      }
      Catch {
        Write-Output $_.Exception.Message
        Write-Host "An error occurred..."
      }
      Finally {
        If ($VMOnline) {
          Write-Host 'Deleting online snapshot...' -fore cyan
          & $VBoxManage snapshot $Global:VMname delete "vrs-temp"
        }
      }

      Try {
        # Decrypt the password in memory
        $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecurePassword)
        $PlainPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
        Remove-Variable BSTR

        # Start the backup
        & $BorgBash -l -c @"
export BORG_PASSPHRASE=${PlainPassword}${EOL}
echo Starting borg backup...${EOL}
cd ${Global:BorgTarget}${EOL}
borg create -vspx -C lz4 ${BackupHost}:${BackupHostPath}::'${Global:BorgArchiveTag}-{now:%Y.%m.%d-%H.%M.%S}-$StateSuffix' .${EOL}
"@
        Remove-Variable PlainPassword
      }
      Catch {
        Write-Output $_.Exception.Message
        Write-Output 'An error has occured. Cleaning up...'
      }
      Finally {
        Write-Host 'Deleting VSS shadow...' -fore cyan
        Send-Tcp-Message 'vss_delete' 8230
        Await-Tcp-Response 8231
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

      $ExtractTarget = ($Global:BorgTarget -replace '\.atomic/' -replace 'cygdrive', 'mnt')

      # Decrypt the password in memory
      $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecurePassword)
      $PlainPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
      Remove-Variable BSTR

      # Mount the borg repository
      Write-Host "Getting list of Archives..."
      $MountpointTest=(& "C:\Program Files\Borg\bin\bash.exe" -l -c "ssh ${BackupHost} 'find ${RemoteMountpoint} -maxdepth 1 ! -path . -printf \`"%f\n\`"'")
      If ($MountpointTest.count -gt 0) {
        & $BorgBash -l -c "ssh ${BackupHost} 'fusermount -u ${RemoteMountpoint}'"
      }

      Write-Host "Verified mountpoint."

      Try {
        & $BorgBash -l -c "ssh ${BackupHost} 'BORG_PASSPHRASE=${PlainPassword} borg mount -o allow_other ${BackupHostPath} ${RemoteMountpoint}'"
        Remove-Variable PlainPassword

        Write-Host "Mount completed"
        $Archives=@(& $BorgBash -l -c "ssh ${BackupHost} 'find ${RemoteMountpoint} -name \`"${Global:BorgArchiveTag}*\`" -maxdepth 1 ! -path . -printf \`"%f\n\`" | sort -n'")

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
          }
          Else {
            Try {
              $ArchiveSelected = $ArchiveSelected -as [int]
            }
            Catch {
              Continue
            }
            If (0 -le $ArchiveSelected -le $Archives.count - 1) {
              $Archive = $Archives[$ArchiveSelected]
              Break
            }
          }
        }
        Write-Host ("Restoring archive {0}..." -f $Archive)

        $Path = ""

        While ($true) {
          $DirScan = @(& $BorgBash -l -c "ssh ${BackupHost} 'find ${RemoteMountpoint}/${Archive}/${Path} -maxdepth 1 ! -path . -printf \`"%f\n\`"'")

          If (($DirScan -like '*.vbox').count -eq 1) {
            Break
          }
          ElseIf ($DirScan.count -gt 1) {
            Throw "Could not detect VM config file in archive!"
          }

          $Path = $Path + ($DirScan[0] -replace " ","\ ") + "/"
        }

        Write-Host ("Detected VM folder: {0}." -f $Path)

        bash -c "time rsync -vah -e 'ssh -Tx -c aes128-gcm@openssh.com' --compress-level=3 --inplace --info=progress2 --stats --delete-before ${BackupHost}:'${RemoteMountpoint}/${Archive}/${Path}' ${ExtractTarget}/"
      }
      Catch {
        Write-Output $_.Exception.Message
        Write-Output "An error has occured. Cleaning up..."
      }
      Finally {
        & $BorgBash -l -c "ssh ${BackupHost} 'fusermount -u ${RemoteMountpoint}'"
        Write-Host "Unmounted filesystem."
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
  } Catch { }
  Finally {
    Send-Tcp-Message "exit" 8230
  }
}
