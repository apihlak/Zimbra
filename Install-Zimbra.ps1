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
               HelpMessage="Winlogbeat 64 bit download path.")]      
       [Alias("EW6")]        
       [ValidateNotNullOrEmpty()]        
       [string[]]        
       $MSIDownload= "https://github.com/apihlak/Sysmon/raw/master/ZimbraConnector8.0.8.1178_x86.msi"      
    )

    Begin {
    }
    Process {
            
        #Install Zimbra
            New-Item -Path $Path_Zimbra -ItemType Directory
            Set-Location -Path "$Path_Zimbra"

            (New-Object System.Net.WebClient).DownloadFile("$MSIDownload","$Path_Zimbra\ZimbraConnector.msi")

            Start-Process msiexec.exe -Wait -ArgumentList '/I ZimbraConnector.msi /lv C:\Zimbra.log /qb'
      }
    End {
        }
    }
