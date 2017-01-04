# We get passed the pwd, so we can look for the config file
param(
  [string]$Pwd
)

# Self-elevating stub, since VSS snapshots can only be created by an administrator
If (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
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

# Query the user for the backup password
$SecurePassword = Read-Host -Prompt "Enter password" -AsSecureString

# Query the user for the backup mode
$Title = "Backup Mode"
$Message = "What backup strategy do you want to use?"
$Automatic = New-Object System.Management.Automation.Host.ChoiceDescription "&Automatic", `
    "Backup automatically every 20 minutes."
$Manual = New-Object System.Management.Automation.Host.ChoiceDescription "&Manual", `
    "Backup manually on user request."
$Options = [System.Management.Automation.Host.ChoiceDescription[]]($Automatic, $Manual)
$Mode = $host.ui.PromptForChoice($Title, $Message, $Options, 0)

switch ($Mode) {
  0 {$Mode = "auto"}
  1 {$Mode = "manual"}
}

# Detect if the system has a battery (so we can avoid killing it)
$Battery = $false
If ($Mode -eq "auto") {
  Try {
    If ((Get-WmiObject -Class BatteryStatus -Namespace root\wmi -List) -ne $null) {
      Write-Host "Detected internal battery"
      $Battery = True
    }
  } 
  Catch { }
}

# Detect the Virtual Machine path
If (Test-Path 'E:\VirtualBox VMs') {
  Write-Host "Detected external VM folder"
  $VSSTarget = "E:\.atomic"
  $VSSDriveLetter = "E"
  $BorgTarget = "/cygdrive/e/.atomic/VirtualBox\\ VMs/"
}
Else {
  Write-Host "Detected internal VM folder"
  $VSSTarget = "C:\.atomic"
  $VSSDriveLetter = "C"
  $BorgTarget = "/cygdrive/c/.atomic/Users/Pytho/VirtualBox\\ VMs/"
}

While ($true) {
  $Counter = 0
  If ($Battery -and !(Get-WmiObject -Class BatteryStatus -Namespace root\wmi).PowerOnLine) {
    Write-Host 'On battery power. Plug laptop in or press any key to backup:'

    While (!$Host.UI.RawUI.KeyAvailable) {
      [Threading.Thread]::Sleep(1000)
      $Counter++
      If ($Counter % 60 * 1 -eq 0) {
        If ((Get-WmiObject -Class BatteryStatus -Namespace root\wmi).PowerOnLine) {
          Break
        }
      }
    }

    # Clear the pending key
    If ($host.ui.RawUI.KeyAvailable) {
      $host.ui.RawUI.ReadKey("IncludeKeyDown,NoEcho")
    }
  }

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
    $StateSuffix = If ($VMsOnline) { "offline" } Else { "online" } 

    # Decrypt the password in memory
    $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecurePassword)
    $PlainPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
    Remove-Variable BSTR

    # Start the backup
    & "C:\Program Files\Borg\bin\bash.exe" -l -c @"
    export BORG_PASSPHRASE=${PlainPassword}${EOL}
    echo Starting borg backup...${EOL}
    cd $BorgTarget${EOL}
    borg create -vspx -C lz4 ${BackupHost}:${BackupHostPath}::'{now:%Y.%m.%d-%H.%M.%S}-vms-$StateSuffix-{hostname}-{user}' .${EOL}
    echo Starting borg prune...${EOL}
    borg prune -vs --list ${BackupHost}:${BackupHostPath} --keep-within 2H -H 8 -d 7 -w 3${EOL}
"@
    Remove-Variable PlainPassword
  }
  Catch {
    Write-Output "An error has occured. Cleaning up..."
  }
  Finally {
    # Remove the VSS target and delete the VSS snapshot
    cmd /c "rmdir /s /q ${VSSTarget}"
    vssadmin delete shadows /shadow=$ShadowID /quiet
    Write-Output "Successfully deleted shadow copy ${ShadowID}."
  }
  
  If ($Mode -eq "manual") {
    Read-Host 'Press Enter to backup again...' | Out-Null
  }
  Else {
    Write-Host 'Waiting 20 minutes til next backup... Press any key to backup immediately:'

    While (!$Host.UI.RawUI.KeyAvailable -and ($Counter++ -lt 60 * 1)) {
      [Threading.Thread]::Sleep(1000)
    }
    # Clear the pending key
    If ($host.ui.RawUI.KeyAvailable) {
      $host.ui.RawUI.ReadKey("IncludeKeyDown,NoEcho")
    }
  }
}
