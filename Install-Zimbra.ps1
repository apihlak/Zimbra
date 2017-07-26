function Install-Zimbra {
<#
.Synopsis
   Install Zimbra Connector remotely

.Example 
   Invoke-Command -ComputerName COMPUTERNAME -ScriptBlock { powershell "IEX (New-Object Net.WebClient).DownloadString('https://raw.githubusercontent.com/apihlak/Zimbra/master/Install-Zimbra.ps1');Install-Zimbra" }

#>
    [CmdletBinding()]
    Param (
       [Parameter(Mandatory=$false,       
               Position=0,       
               HelpMessage="Zimbra download path.")]     
       [Alias("PS")]     
       [ValidateNotNullOrEmpty()]        
       [string[]]        
       $Path_Zimbra = "$env:SystemDrive\ProgramData\ZimbraConnector",     
              
       [Parameter(Mandatory=$false,      
               Position=1,       
               HelpMessage="Zimbra 32 bit download path.")]      
       [Alias("EW6")]        
       [ValidateNotNullOrEmpty()]        
       [string[]]        
       $MSIDownload= "https://github.com/apihlak/Zimbra/raw/master/ZimbraConnector8.0.8.1178_x86.msi"      
    )

    Begin {
    }
    Process {
            
    If ((Get-ItemProperty "HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*" | Select-Object DisplayName | Where-Object { $_.DisplayName -Like '*Zimbra*' }) -eq $null) {    

            #Make directory for Zimbra msi
            New-Item -Path $Path_Zimbra -ItemType Directory   

            #Set location
            Set-Location -Path "$Path_Zimbra"   

            #Install file
            (New-Object System.Net.WebClient).DownloadFile("$MSIDownload","$Path_Zimbra\ZimbraConnector.msi")
            
            #Install Zimbra Connector
            Start-Process msiexec.exe -Wait -ArgumentList '/I ZimbraConnector.msi /qb'
      }
    }  
    End {
        }
    }
