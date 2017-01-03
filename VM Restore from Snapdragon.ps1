# Query the user for the backup password
$SecurePassword = Read-Host -Prompt "Enter password" -AsSecureString

$ExtractRepo = "/mnt/backup/Borgs/VMsBorg"
$ExtractSource = "/mnt/backup/Borgs/mnt"
$Location = "pythonnut@snapdragon"

# Detect the Virtual Machine path
If (Test-Path 'E:\VirtualBox VMs') {
  Write-Host "Detected external VM folder"
  $ExtractTarget = "/mnt/e/VirtualBox\\ VMs/"
}
Else {
  Write-Host "Detected internal VM folder"
  $ExtractTarget = "/mnt/c//Users/Pytho/VirtualBox\\ VMs/"
}

# Decrypt the password in memory
$BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecurePassword)
$PlainPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
Remove-Variable BSTR

# Mount the borg repository
Try {
  Write-Host "Getting list of Archives..."
  $MountpointTest=(& "C:\Program Files\Borg\bin\bash.exe" -l -c "ssh $Location 'ls -1 $ExtractSource'")
  If ($MountpointTest.count -gt 0) {
    & "C:\Program Files\Borg\bin\bash.exe" -l -c "ssh $Location 'fusermount -u $ExtractSource'"
  }
  Write-Host "Verified mountpoint"
  & "C:\Program Files\Borg\bin\bash.exe" -l -c "ssh $Location 'BORG_PASSPHRASE=${PlainPassword} borg mount -o allow_other $ExtractRepo $ExtractSource'"
  Write-Host "Mount completed"
  $Archives=(& "C:\Program Files\Borg\bin\bash.exe" -l -c "ssh $Location 'ls -1 $ExtractSource'")
  Write-Host "Archive list retrieved"
  Write-Host ""
  Write-Host "Select archive to restore:"
  For ($I = 0; $I -lt $Archives.count; $I++) {
    Write-Host ("[{0:00}] {1}" -f $I, $Archives[$I])
  }

  $Archive = $Archives[(Read-Host 'Enter archive number') -as [int]]
  Write-Host ("Restoring archive {0}..." -f $Archive)

  $Path = ""

  While ($true) {
    $DirScan = (& "C:\Program Files\Borg\bin\bash.exe" -l -c "ssh $Location 'ls -A1 ${ExtractSource}/${Archive}/${Path}'")
    If ($DirScan -eq $null) {
      Throw "Could not detect VM folder in archive"
    }

    If ($DirScan -contains "Arch Linux Aurora") {
      Break
    }

    If ($DirScan.GetType() -ne "".GetType()) {
      Throw "Could not detect VM folder in archive"
    }
    
    $Path = $Path + ($DirScan -replace " ","\ ") + "/"
  }

  Write-Host ("Detected VM folder: {0}." -f $Path)

  bash -c "time rsync -vah -e 'ssh -Tx -c aes128-gcm@openssh.com -o Compression=no' --inplace --info=progress2 --stats --delete-before pythonnut@192.168.0.117:'${ExtractSource}/${Archive}/${Path}' ${ExtractTarget}"
}
Catch {
  Write-Output $_.Exception.Message
  Write-Output "An error has occured. Cleaning up..."
}
Finally {
  & "C:\Program Files\Borg\bin\bash.exe" -l -c "ssh $Location 'fusermount -u $ExtractSource'"
  Write-Host "Unmounted filesystem"
  Read-Host 'Press Enter to exit...' | Out-Null
}
