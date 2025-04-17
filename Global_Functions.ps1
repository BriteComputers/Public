Function Set-NoSleep{

Write-Host "Set Laptop not to sleep while plugged in"

Powercfg /Change monitor-timeout-ac 0
Powercfg /Change monitor-timeout-dc 10
Powercfg /Change standby-timeout-ac 0
Powercfg /Change standby-timeout-dc 30

}
Function Set-ESTTime{

Write-Host "Setting to Eastern Time Zones"

Set-TimeZone -Name "Eastern Standard Time"
net start W32Time
W32tm /resync /force

}
Function Disable-FastStartup {
    Write-Host "Disable Windows Fast Startup"
        REG ADD "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Session Manager\Power" /v HiberbootEnabled /t REG_DWORD /d "0" /f
        powercfg -h off
}
Function Update-Windows{
    param
    (
        [ValidateSet('Yes','No')]
        [Parameter(Mandatory=$false)]
        [string]$HideUpdates
    )

    $Folder = 'C:\IT\Update-Logs'
    if (Test-Path -Path $Folder) {
        "Folder Exists"
    } else {
        mkdir C:\IT\Update-Logs
    }

    $progressPreference = 'silentlyContinue'

    # Installs NuGet with Forced
    Get-PackageProvider -Name "nuGet" -ForceBootstrap | 
        Select-Object -Property Name, Version | 
        Format-Table -Autosize

    # Trusts Microsofts PSGallery
    Set-PSRepository -Name 'PSGallery' -InstallationPolicy Trusted
    Get-PSRepository -Name 'PSGallery' | Format-List * -Force
        
    # Install PSWindowsUpdate Module
    Install-Module PSWindowsUpdate

    $WUBlocklist = "KB5053598"
    if ($HideUpdates -eq "Yes"){
        foreach ($element in $WUBlocklist) {
            Hide-WindowsUpdate -KBArticleID  $element -AcceptAll
            Write-Host "Hid windows update $element Temporarily"
        }
    }
    Else {
        Show-WindowsUpdate -AcceptAll
    }

    Get-WindowsUpdate | Out-File C:\IT\Update-Logs\Updates_"$((Get-Date).ToString('dd-MM-yyyy_HH.mm.ss'))".txt

    $Count = 0
    $Attempts = 15
    $ProgressPreference = 'SilentlyContinue'
    while( $true ){
        $Error.Clear()
        $Count++
        Write-Host "Attempt #" $Count "to update Windows"
        try{

            If($Count -lt $Attempts){
                Install-WindowsUpdate -AcceptAll -IgnoreReboot
                Break
            }

            Else{
                
                Write-Host "Updates exceeded 15 attempts"
                Break
            }
        }

        catch{
            Write-Host "Error.... Retrying"
            Start-Sleep -Seconds 1
        }
    }
}
Function Set-AutoLogon {
    
    param
    (
        [Parameter(Mandatory=$true)]
        [string]$Username,
        [Parameter(Mandatory=$true)]
        [string]$Password
    )
    
    #$Password = $Password | ConvertTo-SecureString -asPlainText -Force
    Write-Host "Set autologon"
    #Registry path declaration
    $RegPath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon"
    [String]$DefaultUsername = $Username
    [String]$DefaultPassword = $Password
    #setting registry values
    Set-ItemProperty $RegPath "AutoAdminLogon" -Value "1" -type String
    Set-ItemProperty $RegPath "DefaultUsername" -Value $DefaultUsername -type String
    Set-ItemProperty $RegPath "DefaultPassword" -Value $DefaultPassword -type String
    Set-ItemProperty $RegPath "AutoLogonCount" -Value "1" -type DWord
    Write-Host "End of Set autologon"
}
Function Join-Domain {
        param
    (
        [Parameter(Mandatory=$true)]
        [string]$Domain,
        [Parameter(Mandatory=$true)]
        [string]$Username,
        [Parameter(Mandatory=$true)]
        $Password
    )
    Write-Host "Join Domain"
    $Password = $Password | ConvertTo-SecureString -asPlainText -Force
    $Username = $Domain + "\" + $Username
    $credential = New-Object System.Management.Automation.PSCredential($Username,$Password)
    Add-Computer -DomainName $Domain -Credential $credential
}
Function Remove-PPKGInstallFolder {

    Write-Host "Cleaning up and Restarting Computer"
    PowerShell.exe -NoProfile -ExecutionPolicy Bypass -Command "If (Test-Path C:\IT\PPKG){Remove-Item -LiteralPath 'C:\IT\PPKG' -Force -Recurse};Restart-Computer -Force"
    Stop-transcript
    Restart-Computer -Force

}
Function Add-WebShortcut{

    param
    (
        [string]$Label,
        [string]$Url
    )

    $Folder = 'C:\Temp\Shortcuts'
    if (-not (Test-Path -Path $Folder)) {
        mkdir C:\Temp\Shortcuts
    }

    Write-Host "Adding a shortcut to $Label to the Temp Folder"
    $Shell = New-Object -ComObject ("WScript.Shell")
    $URLFilePath = "$Folder\" + $Label + ".url"
    $Favorite = $Shell.CreateShortcut($URLFilePath)
    $Favorite.TargetPath = $Url
    $Favorite.Save()

}
Function Set-PCName{

    # Sample script to rename a domain joined computer
    $SerialNumber = (Get-WmiObject win32_bios).SerialNumber
    $NewComputerName = $SiteCode + "-" + $SerialNumber

    Write-Host "Rename Computer to" $NewComputerName
    Rename-Computer -NewName $NewComputerName

}
Function Set-ProtectionPolicy{

    Write-Host "Setting Protection Policy REG Key to Fix 365 app Login Errors"
    Set-ItemProperty 'HKLM:\Software\Microsoft\Cryptography\Protect\Providers\df9d8cd0-1501-11d1-8c7a-00c04fc297eb' -Name "ProtectionPolicy" -Value "1" -type DWord

}
Function Update-WindowTitle ([String] $PassNumber) {
    Write-Host "Changing window title"
    $host.ui.RawUI.WindowTitle = "Provisioning | $env:computername | Pass $PassNumber | Please Wait"
}
Function Start-PPKGLog ([String] $LogLabel) {
    Write-Host "Making a log file for debugging"
        $LogPath = "C:\IT\" + $LogLabel + ".log"
        Start-Transcript -path $LogPath -Force -Append
}
Function Install-Apps{
    Write-Host "Installing Apps"

    ##Install Adobe and Chrome for all sites
    Install-Adobe
    Install-Chrome
    Install-O365
    
    Switch ($SiteCode){
        "122" {
            Install-Teams
        }

        "312" {
            Install-Teams
            Install-Anyconnect
        }

        "114" {
            Install-Teams
        }

        "109" {
            Install-Teams
        }

        "287" {
            Install-Teams
            Install-7Zip
        }
        "374" {
        }

        default {
            Write-Host "Site not set up in Global Functions"
        }
    }
}
Function Set-RunOnce{

    param
    (
        [string]$Label
    )

    $RunOnceValue = 'PowerShell.exe -ExecutionPolicy Bypass -File "C:\IT\PPKG\' + $Label + '.ps1"'
    Write-Host "Install After Reboot"
    Set-ItemProperty 'HKLM:\Software\Microsoft\Windows\CurrentVersion\RunOnce' -Name $Label -Value $RunOnceValue
    
}
Function Connect-Wifi {
    param
        (
            [Parameter(Mandatory=$False)]
            [string]$NetworkSSID,

            [Parameter(Mandatory=$true)]
            [string]$NetworkPassword,

            [ValidateSet('WEP','WPA','WPA2','WPA2PSK')]
            [Parameter(Mandatory=$False)]
            [string]$Authentication = 'WPA2PSK',

            [ValidateSet('AES','TKIP')]
            [Parameter(Mandatory=$False)]
            [string]$Encryption = 'AES'
        )

    # Create the WiFi profile, set the profile to auto connect
    $WirelessProfile = @'
<WLANProfile xmlns="http://www.microsoft.com/networking/WLAN/profile/v1">
    <name>{0}</name>
    <SSIDConfig>
        <SSID>
            <name>{0}</name>
        </SSID>
    </SSIDConfig>
    <connectionType>ESS</connectionType>
    <connectionMode>auto</connectionMode>
    <MSM>
        <security>
            <authEncryption>
                <authentication>{2}</authentication>
                <encryption>{3}</encryption>
                <useOneX>false</useOneX>
            </authEncryption>
            <sharedKey>
                <keyType>passPhrase</keyType>
                <protected>false</protected>
                <keyMaterial>{1}</keyMaterial>
            </sharedKey>
        </security>
    </MSM>
</WLANProfile>
'@ -f $NetworkSSID, $NetworkPassword, $Authentication, $Encryption
        
    
    # Create the XML file locally
    $random = Get-Random -Minimum 1111 -Maximum 99999999
    $tempProfileXML = "$env:TEMP\tempProfile$random.xml"
    $WirelessProfile | Out-File $tempProfileXML

    # Add the WiFi profile and connect
    Start-Process netsh ('wlan add profile filename={0}' -f $tempProfileXML)

    # Connect to the WiFi network - only if you need to
    $WifiNetworks = (netsh wlan show network)
    $NetworkSSIDSearch = '*' + $NetworkSSID + '*'
    If ($WifiNetworks -like $NetworkSSIDSearch) {
        Write-Host "Found SSID: $NetworkSSID `nAttempting to connect"
        Start-Process netsh ('wlan connect name="{0}"' -f $NetworkSSID)
        Start-Sleep 5
        netsh interface show interface
    } Else {
        Write-Host "Did not find SSID: $NetworkSSID `nConnection profile stored for later use."
    }
}
Function Install-O365{
    
    $DowloadURL = "https://download.microsoft.com/download/2/7/A/27AF1BE6-DD20-4CB4-B154-EBAB8A7D4A7E/officedeploymenttool_16731-20398.exe"
    $TempPath = "C:\IT\O365"
    $DownloadFile = "$TempPath\O365-Installer.exe"

    if (!(Test-Path $TempPath)) {
        New-Item -ItemType "Directory" -Path $TempPath
    }
    
    Write-Host "Downloading 365 Apps"
    $ProgressPreference = 'SilentlyContinue'
    Invoke-WebRequest $DowloadURL -OutFile $DownloadFile
    Write-Host "Extracting Setup.exe file"
    Start-Process $DownloadFile -ArgumentList "/quiet /extract:$temppath" -wait

    $O365ConfigDest = "C:\IT\O365\configuration-Office365-x64.xml"
    Write-Host "Installing Office"
    & C:\IT\O365\setup.exe /configure $O365ConfigDest | Wait-Process
    Write-Host "Placing Shortcuts"
        If (Test-Path "C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE"){
            $TargetFile = "C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE"
        } ELSEIF (Test-Path "C:\Program Files (x86)\Microsoft Office\root\Office16\OUTLOOK.EXE"){
            $TargetFile = "C:\Program Files (x86)\Microsoft Office\root\Office16\OUTLOOK.EXE"
        }
        $ShortcutFile = "$env:Public\Desktop\Outlook.lnk"
        $WScriptShell = New-Object -ComObject WScript.Shell
        $Shortcut = $WScriptShell.CreateShortcut($ShortcutFile)
        $Shortcut.TargetPath = $TargetFile
        $Shortcut.Save()

        If (Test-Path "C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE"){
            $TargetFile = "C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE"
        } ELSEIF (Test-Path "C:\Program Files (x86)\Microsoft Office\root\Office16\EXCEL.EXE"){
            $TargetFile = "C:\Program Files (x86)\Microsoft Office\root\Office16\EXCEL.EXE"
        }
        $ShortcutFile = "$env:Public\Desktop\Excel.lnk"
        $WScriptShell = New-Object -ComObject WScript.Shell
        $Shortcut = $WScriptShell.CreateShortcut($ShortcutFile)
        $Shortcut.TargetPath = $TargetFile
        $Shortcut.Save()

        If (Test-Path "C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE"){
            $TargetFile = "C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE"
        } ELSEIF (Test-Path "C:\Program Files (x86)\Microsoft Office\root\Office16\WINWORD.EXE"){
            $TargetFile = "C:\Program Files (x86)\Microsoft Office\root\Office16\WINWORD.EXE"
        }
        $ShortcutFile = "$env:Public\Desktop\Word.lnk"
        $WScriptShell = New-Object -ComObject WScript.Shell
        $Shortcut = $WScriptShell.CreateShortcut($ShortcutFile)
        $Shortcut.TargetPath = $TargetFile
        $Shortcut.Save()
}
Function Connect-VPN{
    param
        (
            [ValidateSet('WindowsVPN','Netextender')]			
            [Parameter(Mandatory=$True)]
            [string]$VPNType,

            [Parameter(Mandatory=$True)]
            [string]$IPAddress,

            [Parameter(Mandatory=$False)]
            [string]$Domain,

            [Parameter(Mandatory=$True)]
            [string]$Username,

            [Parameter(Mandatory=$True)]
            [string]$Password,

            [Parameter(Mandatory=$False)]
            [string]$Name
        )

    If($VPNType -eq "WindowsVPN"){
        Start-Process rasdial -NoNewWindow -ArgumentList "$Name $Username $Password" -PassThru -Wait
    }
    IF($VPNType -eq "Netextender"){

        Start-Process -FilePath "C:\Program Files (x86)\SonicWall\SSL-VPN\NetExtender\NECLI.exe" -ArgumentList "connect -s $IPAddress -d $Domain -u $Username -p $Password --always-trust"
    }
    Else{
        Write-Host "VPN connection not set up yet, see your Autodeploy Master"
    }
}
function Set-AllowPing{
    
    netsh advfirewall firewall add rule name="ICMP Allow incoming V4 echo request" protocol=icmpv4:8,any dir=in action=allow
}
Function Enable-NetBios{

    $i = 'HKLM:\SYSTEM\CurrentControlSet\Services\netbt\Parameters\interfaces'  
    Get-ChildItem $i | ForEach-Object {  
        Set-ItemProperty -Path "$i\$($_.pschildname)" -name NetBiosOptions -value 1
    }
}
Function Disable-IPV6{

    Disable-NetAdapterBinding -Name "*" -ComponentID ms_tcpip6

}
Function Install-Adobe{

    Install-StoreApp -PackageID "XPDP273C0XHQH2" -Log "AdobeWingetInstall.log"
    
}
Function Install-AdobePro{

    $directory = "C:\IT\Apps"

    If ((Test-Path -Path $directory) -eq $false)
    
        {
    
            New-Item -Path $directory -ItemType directory
    
        }
    
    Write-Host "Adobe is not installed. Downloading latest version..."
    Set-Location $directory
    $ProgressPreference = 'SilentlyContinue'
    $downloadPath = "$directory\Adobe-Installer.zip"
    Invoke-WebRequest "https://trials.adobe.com/AdobeProducts/APRO/Acrobat_HelpX/win32/Acrobat_DC_Web_x64_WWMUI.zip" -OutFile $downloadPath
    Write-Host "Extracting installer"
    Expand-Archive -path $downloadPath -DestinationPath $directory
    Start-Sleep -Seconds 30
    Write-Host "Installing Adobe Pro"
    
    If(Test-Path "$directory\Adobe Acrobat\Setup.exe"){
    
        Start-Process "$directory\Adobe Acrobat\Setup.exe" -ArgumentList "/sl '1033' /sALL" -Wait
        Write-Host "Adobe Pro Installed"
    
        }
    Else{
        
        Write-Host "Adobe Failed to Install"
    
        }
}
function Install-Agent {
    Param(
        $Token,
        $Domain
    )
    $TempPath = "C:\IT\Agent"
    $DownloadPath = "$TempPath\WindowsAgentSetup.exe"
    $AgentDownload = "https://rmm.$Domain/download/2024.6.0.19/winnt/N-central/WindowsAgentSetup.exe"

    if (!(Test-Path $TempPath)) {
        New-Item -ItemType "Directory" -Path $TempPath
    }
    
    $progressPreference = 'silentlyContinue'
    Invoke-Webrequest $AgentDownload -OutFile $DownloadPath
    
    Start-Process $DownloadPath -ArgumentList "/s /v"" /qn CUSTOMERID=$SiteCode REGISTRATION_TOKEN=$Token CUSTOMERSPECIFIC=1 SERVERPROTOCOL=HTTPS SERVERADDRESS=rmm.$Domain SERVERPORT=443""" -wait

}
Function Install-RingCentral{
    
    Write-Host "RingCentral is not installed. Downloading latest version..."
    mkdir  C:\RC-Temp
    Set-Location C:\RC-Temp
    $ProgressPreference = 'SilentlyContinue'
    Invoke-WebRequest "https://app.ringcentral.com/download/RingCentral.exe" -OutFile C:\RC-Temp\RingCentral.exe
    Write-Host "RingCentral is downloaded. Installing..."
    ./RingCentral.exe /S | Wait-Process
    Write-Host "Waiting 5 sec before removing temp files..."
    Start-Sleep -Seconds 5
    Set-Location ..
    Remove-Item -Force -Recurse C:\RC-Temp 
    
    Start-Sleep -Seconds 60
    TASKKILL /F /IM RingCentral.exe
    
}	
Function Install-Zoom{
    
    $DownloadPath = "https://zoom.us/client/latest/ZoomInstallerFull.msi?archType=x64"
    $ZOPDownloadPath = "https://zoom.us/client/latest/ZoomOutlookPluginSetup.msi"
    $TempPath = "C:\Zoom-Temp"
    $DownloadFile = "$TempPath\Zoom-Installer.msi"
    $ZOPDownloadFile = "$TempPath\ZoomOutlookPlugin-Installer.msi"
    
    Write-Host "Zoom is not installed. Downloading latest version..."
    mkdir  $TempPath
    Set-Location $TempPath
    $ProgressPreference = 'SilentlyContinue'
    Invoke-WebRequest $DownloadPath -OutFile $DownloadFile
    Write-Host "Zoom is downloaded. Installing..."
    Start-Process msiexec.exe -ArgumentList "/i Zoom-Installer.msi /quiet" -wait
    Write-Host "Waiting 5 sec before removing temp files..."
    Start-Sleep -Seconds 5
    
    Write-Host "Zoom outlook Plugin Downloading latest version..."
    Invoke-WebRequest $ZOPDownloadPath -OutFile $ZOPDownloadFile
    Start-Process msiexec.exe -ArgumentList "/i ZoomOutlookPlugin-Installer.msi /quiet" -wait
    Write-Host "Waiting 5 sec before removing temp files..."
    Start-Sleep -Seconds 5
    
    Set-Location ..
    Remove-Item -Force -Recurse $TempPath
    
}
Function Install-Vantage{

    Write-Host "Installing Lenovo Commercial Vantage from MS Store"
    Install-Module -Name Microsoft.WinGet.Client
    Add-WUServiceManager -ServiceID 117cab2d-82b1-4b5a-a08c-4d62dbee7782 -Confirm:$false -Verbose
    winget install "Commercial Vantage" --source msstore --Accept-source-agreements --accept-package-agreements
    Write-Host "Lenovo Commercial Vantage"

}
function Install-Teams{

    $TempDir = "C:\IT\Apps"
    $DownloadPath = "$TempDir\TeamsInstaller.msi"

    Write-Host "Downloading Teams"
    $ProgressPreference = 'SilentlyContinue'
    Invoke-WebRequest "https://teams.microsoft.com/downloads/desktopurl?env=production&plat=windows&arch=x64&managedInstaller=true&download=true" -OutFile $DownloadPath
    Write-Host "Installing Teams"
    Start-Process MsiExec.exe -ArgumentList "/i $DownloadPath /qn /norestart" -Wait
    Write-Host "Teams Installed"
    
}
Function Uninstall-Bloat {

    $Bloatware = @(

        #Unnecessary Windows 10 AppX Apps
        "Microsoft.BingNews"
        "Microsoft.GetHelp"
        "Microsoft.Getstarted"
        "Microsoft.Messaging"
        "Microsoft.Microsoft3DViewer"
        "Microsoft.MicrosoftOfficeHub"
        "Microsoft.MicrosoftSolitaireCollection"
        "Microsoft.NetworkSpeedTest"
        "Microsoft.News"
        "Microsoft.Office.Lens"
        "Microsoft.Office.OneNote"
        "Microsoft.Office.Sway"
        "Microsoft.OneConnect"
        "Microsoft.People"
        "Microsoft.Print3D"
        "Microsoft.RemoteDesktop"
        "Microsoft.SkypeApp"
        "Microsoft.StorePurchaseApp"
        "Microsoft.Office.Todo.List"
        "Microsoft.Whiteboard"
        "Microsoft.WindowsAlarms"
        #"Microsoft.WindowsCamera"
        "microsoft.windowscommunicationsapps"
        "Microsoft.WindowsFeedbackHub"
        "Microsoft.WindowsMaps"
        "Microsoft.WindowsSoundRecorder"
        "Microsoft.Xbox.TCUI"
        "Microsoft.XboxApp"
        "Microsoft.XboxGameOverlay"
        "Microsoft.XboxGamingOverlay"
        "Microsoft.XboxIdentityProvider"
        "Microsoft.XboxSpeechToTextOverlay"
        "Microsoft.ZuneMusic"
        "Microsoft.ZuneVideo"

        #Sponsored Windows 10 AppX Apps
        #Add sponsored/featured apps to remove in the "*AppName*" format
        "*EclipseManager*"
        "*ActiproSoftwareLLC*"
        "*AdobeSystemsIncorporated.AdobePhotoshopExpress*"
        "*7Ziplingo-LearnLanguagesforFree*"
        "*PandoraMediaInc*"
        "*CandyCrush*"
        "*BubbleWitch3Saga*"
        "*Wunderlist*"
        "*Flipboard*"
        "*Twitter*"
        "*Facebook*"
        "*Spotify*"
        "*Minecraft*"
        "*Royal Revolt*"
        "*Sway*"
        "*Speed Test*"
        "*Dolby*"
                
        #Optional: Typically not removed but you can if you need to for some reason
        #"*Microsoft.Advertising.Xaml_10.1712.5.0_x64__8wekyb3d8bbwe*"
        #"*Microsoft.Advertising.Xaml_10.1712.5.0_x86__8wekyb3d8bbwe*"
        #"*Microsoft.BingWeather*"
        #"*Microsoft.MSPaint*"
        #"*Microsoft.MicrosoftStickyNotes*"
        #"*Microsoft.Windows.Photos*"
        #"*Microsoft.WindowsCalculator*"
        #"*Microsoft.WindowsStore*"
    )
    foreach ($Bloat in $Bloatware) {
        Get-AppxPackage -Name $Bloat| Remove-AppxPackage
        Get-AppxProvisionedPackage -Online | Where-Object DisplayName -like $Bloat | Remove-AppxProvisionedPackage -Online
        Write-Output "Trying to remove $Bloat."
    }
}
Function Install-7Zip{

    Write-Host "7Zip is not installed. Downloading latest version..."
    mkdir  C:\7Zip-tmp
    Set-Location C:\7Zip-tmp
    $ProgressPreference = 'SilentlyContinue'
    Invoke-WebRequest "https://www.7-zip.org/a/7z2301-x64.exe" -OutFile C:\7Zip-tmp\7Zip.exe
    Write-Host "7Zip is downloaded. Installing..."
    Start-Process "C:\7Zip-tmp\7Zip.exe" -ArgumentList "/S" -wait
    Write-Host "Waiting 5 sec before removing temp files..."
    Start-Sleep -Seconds 5
    Set-Location ..
    Remove-Item -Force -Recurse C:\7Zip-tmp 
    Write-Host "All files cleaned up. Exiting..."
    
}
function Install-WinGet{
    $tempFolderName = 'WinGetInstall'
    $tempFolder = Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath $tempFolderName
    New-Item $tempFolder -ItemType Directory -ErrorAction SilentlyContinue | Out-Null
    
    $apiLatestUrl = if ($Prerelease) { 'https://api.github.com/repos/microsoft/winget-cli/releases?per_page=1' }
    else { 'https://api.github.com/repos/microsoft/winget-cli/releases/latest' }
    
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    $WebClient = New-Object System.Net.WebClient
    
    function Get-LatestUrl
    {
        ((Invoke-WebRequest $apiLatestUrl -UseBasicParsing | ConvertFrom-Json).assets | Where-Object { $_.name -match '^Microsoft.DesktopAppInstaller_8wekyb3d8bbwe.msixbundle$' }).browser_download_url
    }
    
    function Get-LatestHash
    {
        $shaUrl = ((Invoke-WebRequest $apiLatestUrl -UseBasicParsing | ConvertFrom-Json).assets | Where-Object { $_.name -match '^Microsoft.DesktopAppInstaller_8wekyb3d8bbwe.txt$' }).browser_download_url
        
        $shaFile = Join-Path -Path $tempFolder -ChildPath 'Microsoft.DesktopAppInstaller_8wekyb3d8bbwe.txt'
        $WebClient.DownloadFile($shaUrl, $shaFile)
        
        Get-Content $shaFile
    }
    
    $desktopAppInstaller = @{
        fileName = 'Microsoft.DesktopAppInstaller_8wekyb3d8bbwe.msixbundle'
        url	     = $(Get-LatestUrl)
        hash	 = $(Get-LatestHash)
    }
    
    $vcLibsUwp = @{
        fileName = 'Microsoft.VCLibs.x64.14.00.Desktop.appx'
        url	     = 'https://aka.ms/Microsoft.VCLibs.x64.14.00.Desktop.appx'
        hash	 = '9BFDE6CFCC530EF073AB4BC9C4817575F63BE1251DD75AAA58CB89299697A569'
    }
    $uiLibsUwp = @{
        fileName = 'Microsoft.UI.Xaml.2.7.zip'
        url	     = 'https://www.nuget.org/api/v2/package/Microsoft.UI.Xaml/2.7.0'
        hash	 = '422FD24B231E87A842C4DAEABC6A335112E0D35B86FAC91F5CE7CF327E36A591'
    }
    
    $dependencies = @($desktopAppInstaller, $vcLibsUwp, $uiLibsUwp)
    
    Write-Host '--> Checking dependencies'
    
    foreach ($dependency in $dependencies)
    {
        $dependency.file = Join-Path -Path $tempFolder -ChildPath $dependency.fileName
        #$dependency.pathInSandbox = (Join-Path -Path $tempFolderName -ChildPath $dependency.fileName)
        
        # Only download if the file does not exist, or its hash does not match.
        if (-Not ((Test-Path -Path $dependency.file -PathType Leaf) -And $dependency.hash -eq $(Get-FileHash $dependency.file).Hash))
        {
            Write-Host @"
    - Downloading:
        $($dependency.url)
"@
            
            try
            {
                $WebClient.DownloadFile($dependency.url, $dependency.file)
            }
            catch
            {
                #Pass the exception as an inner exception
                throw [System.Net.WebException]::new("Error downloading $($dependency.url).", $_.Exception)
            }
            if (-not ($dependency.hash -eq $(Get-FileHash $dependency.file).Hash))
            {
                throw [System.Activities.VersionMismatchException]::new('Dependency hash does not match the downloaded file')
            }
        }
    }
    
    # Extract Microsoft.UI.Xaml from zip (if freshly downloaded).
    # This is a workaround until https://github.com/microsoft/winget-cli/issues/1861 is resolved.
    
    if (-Not (Test-Path (Join-Path -Path $tempFolder -ChildPath \Microsoft.UI.Xaml.2.7\tools\AppX\x64\Release\Microsoft.UI.Xaml.2.7.appx)))
    {
        Expand-Archive -Path $uiLibsUwp.file -DestinationPath ($tempFolder + '\Microsoft.UI.Xaml.2.7') -Force
    }
    $uiLibsUwp.file = (Join-Path -Path $tempFolder -ChildPath \Microsoft.UI.Xaml.2.7\tools\AppX\x64\Release\Microsoft.UI.Xaml.2.7.appx)
    Add-AppxPackage -Path $($desktopAppInstaller.file) -DependencyPath $($vcLibsUwp.file), $($uiLibsUwp.file)
    # Clean up files
    Remove-Item $tempFolder -recurse -force
}
Function Install-StoreApp{
    param (
    $PackageID,
    $Log
    )

    Start-Transcript -Path "C:\logs\$Log"
    #Test if Winget is installed. If not, try and install it. 
    try {
        WinGet | Out-Null
    }
    catch {
        Install-WinGet
    }
    try {
        Winget | Out-Null
    }
    Catch {
        Write-Host "Winget not found after attempting to install. Stopping operation"
        exit 1
    }

    Winget install --id $PackageID --source msstore --silent --accept-package-agreements --accept-source-agreements 
    Stop-Transcript
}
Function Install-Chrome{

    $DowloadURL = "https://dl.google.com/tag/s/appguid%3D%7B8A69D345-D564-463C-AFF1-A69D9E530F96%7D%26iid%3D%7B925257FD-0F47-8C66-AA6E-E90133DEB98F%7D%26lang%3Den%26browser%3D3%26usagestats%3D0%26appname%3DGoogle%2520Chrome%26needsadmin%3Dtrue%26ap%3Dx64-stable-statsdef_0%26brand%3DGCGB/dl/chrome/install/googlechromestandaloneenterprise64.msi"
    $TempPath = "C:\IT\Apps\Chrome"
    $DownloadFile = "$TempPath\Chrome-Installer.msi"

    if (!(Test-Path $TempPath)) {
        New-Item -ItemType "Directory" -Path $TempPath
    }
    
    Write-Host "Downloading Chrome"
    $ProgressPreference = 'SilentlyContinue'
    Invoke-WebRequest $DowloadURL -OutFile $DownloadFile
    Write-Host "Installing Chrome"
    Start-Process msiexec.exe -ArgumentList "/i $DownloadFile /quiet" -wait
    
}
Function Install-Program{
    param
        (			
            [Parameter(Mandatory=$True)]
            [string]$AppName,

            [Parameter(Mandatory=$True)]
            [string]$DowloadURL,

            [ValidateSet('msi','exe')]
            [Parameter(Mandatory=$True)]
            [string]$FileType
        )

    $TempPath = "C:\IT\Apps\$AppName"
    $DownloadFile = "$TempPath\$AppName-Installer.$FileType"

    if (!(Test-Path $TempPath)) {
        New-Item -ItemType "Directory" -Path $TempPath
    }

    Write-Host "Downloading $AppName"
    $ProgressPreference = 'SilentlyContinue'
    Invoke-WebRequest $DowloadURL -OutFile $DownloadFile
    Write-Host "Installing $AppName"
    if ($FileType -eq "exe") {
        Start-Process $DownloadFile -ArgumentList "/quiet"
    }
    Else{
        Start-Process msiexec.exe -ArgumentList "/i $DownloadFile /qn /passive /norestart" -wait
    }
}

Function Install-Netextender{

    Param(
    $IPAddress,
    $Domain
    )

    Write-Host "Netextender is not installed. Downloading latest version..."
    mkdir  C:\Netextender-temp
    Set-Location C:\Netextender-temp
    Invoke-WebRequest "https://software.sonicwall.com/NetExtender/NetExtender-x64-10.3.1.msi" -OutFile C:\Netextender-temp\Netextender.msi
    Write-Host "Netextender is downloaded. Installing..."
    Start-Process MsiExec.exe -ArgumentList "/i NetExtender.MSI REBOOT=ReallySuppress SERVER=$IPAddress DOMAIN=$Domain /qn" -wait
    Write-Host "Waiting 5 sec before removing temp files..."
    Start-Sleep -Seconds 5
    Set-Location ..
    Remove-Item -Force -Recurse C:\Netextender-temp

}
