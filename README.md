##IN DOMAIN

$name = Get-Adcomputer -Filter {(enabled -eq $True) -and (Name -like '*EE') } -SearchBase "ou=Linxtelecom,dc=office,dc=linxtelecom,dc=com" | Select-Object -ExpandProperty Name
ForEach ($computer in $name) {
  if (Test-WSMan -ComputerName $computer -ErrorAction Ignore) {
  Write-Warning -Message “Connecting to $computer”
  Invoke-Command -ComputerName $computer -ScriptBlock { 
      If ((Get-ItemProperty "HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*" | Select-Object DisplayName | Where-Object { $_.DisplayName -Like '*Zimbra*' }) -eq $null) {    
        Write-Warning -Message “Running script”
        powershell "IEX (New-Object Net.WebClient).DownloadString('https://raw.githubusercontent.com/apihlak/Zimbra/master/Install-Zimbra.ps1');Install-Zimbra"
      } else {
        Write-Warning -Message “Zimbra Connector exists”
      }
    }
  } else {
  Write-Warning -Message “Couldn’t connect to $computer”
  }
}

##SEPARATE
powershell "IEX (New-Object Net.WebClient).DownloadString('https://raw.githubusercontent.com/apihlak/Zimbra/master/Install-Zimbra.ps1');Install-Zimbra"
