#!C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe
<#
	2019 powershell script for the Prep of Vue Pacs servers
		PTD	June 2022
			modified 9-19-2022

#>	
#
# check all
    netsh advfirewall show allprofiles
    Get-NetFirewallRule -DisplayGroup "Remote Desktop"| select-Object Enabled, Profile
    Get-NetFirewallRule -DisplayGroup "Remote Desktop (WebSocket)"| select-Object Enabled, Profile
    Get-NetFirewallRule -displayName "VNC Server"| select-Object Enabled, Profile
    Get-WindowsFeature | findstr "Multipath Server-Backup Telnet-Client"
    ls "HKLM:\SOFTWARE\Microsoft\ServerManager*"
    ls "HKLM:\SYSTEM\*ControlSet*\Control\Terminal Server*"| findstr Temp
    ls "HKLM:\SYSTEM\*ControlSet*\Control\Session Manager\SubSystems*" |findstr Shared
    netsh int ipv4 show dynamicportrange protocol=tcp
    ls "HKCU:\Software\Microsoft\Windows\DWM*" |findstr ColorPrevalence
    ls "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced*" |findstr "Hidden Always HideFileExt LaunchTo"
    ls "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\HideDesktopIcons\NewStartPanel*"
    ls "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\policies\system*" |findstr Prompt
    ls "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer"|findstr Visual
    wmic OS Get DataExecutionPrevention_SupportPolicy
    Gwmi win32_Pagefilesetting | Select Name, InitialSize, MaximumSize
    Get-TimeZone
#
read-host "check all the current setting and hit enter to continue"
#
echo "Calculating how much RAM this system has"
    $ram = ((Get-CimInstance Win32_PhysicalMemory | Measure-Object -Property capacity -Sum).sum /1mb)
    $tz = (read-host "what time zone is this system in?, enter just the 1st part name, like Eastern, Central, Mountain, Pacific, etc")
#
# turn off firewalls
    netsh advfirewall set allprofiles state off
    Set-NetFirewallRule -DisplayGroup "Remote Desktop" -profile "Domain, Private, Public" -enable True
    Set-NetFirewallRule -DisplayGroup "Remote Desktop (WebSocket)" -profile "Domain, Private, Public" -enable True
    Set-NetFirewallRule -DisplayName "VNC Server" -profile "Domain, Private, Public" -enable True
#
# set tcp ports
    netsh int ipv4 set dynamicportrange protocol=tcp startport=45152 numberofports=20384
#
# set server manager not to auto launch
    Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\ServerManager" -Name "DoNotOpenServerManagerAtLogon" -value 1
#
# RDP settings
    Set-ItemProperty -Path "HKLM:\SYSTEM\ControlSet001\Control\Terminal Server" -Name "DeleteTempDirsOnExit" -Value 0
    Set-ItemProperty -Path "HKLM:\SYSTEM\ControlSet001\Control\Terminal Server" -Name "PerSessionTempDir" -Value 0
    Set-ItemProperty -Path "HKLM:\SYSTEM\ControlSet002\Control\Terminal Server" -Name "DeleteTempDirsOnExit" -Value 0
    Set-ItemProperty -Path "HKLM:\SYSTEM\ControlSet002\Control\Terminal Server" -Name "PerSessionTempDir" -Value 0
#
# desktop heap parameters and open sockets
    Set-ItemProperty -Path "HKLM:\SYSTEM\ControlSet001\Control\Session Manager\SubSystems" -name Windows -value "%SystemRoot%\system32\csrss.exe ObjectDirectory=\Windows SharedSection=1024,20480,6144 Windows=On SubSystemType=Windows ServerDll=basesrv,1 ServerDll=winsrv:UserServerDllInitialization,3 ServerDll=sxssrv,4 ProfileControl=Off MaxRequestThreads=16"
    Set-ItemProperty -Path "HKLM:\SYSTEM\ControlSet002\Control\Session Manager\SubSystems" -name Windows -value "%SystemRoot%\system32\csrss.exe ObjectDirectory=\Windows SharedSection=1024,20480,6144 Windows=On SubSystemType=Windows ServerDll=basesrv,1 ServerDll=winsrv:UserServerDllInitialization,3 ServerDll=sxssrv,4 ProfileControl=Off MaxRequestThreads=16"
#
# set color and explorer folder preferences
    Set-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Windows\DWM" -Name "ColorPrevalence" -value 1
    Set-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -Name "HideFileExt" -value 0
    Set-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -Name "Hidden" -value 1
    Set-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -Name "LaunchTo" -value 1
    Set-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -Name "AlwaysShowMenus" -value 1
    Set-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\HideDesktopIcons\NewStartPanel" -Name "{20D04FE0-3AEA-1069-A2D8-08002B30309D}" -value 0
    Set-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\HideDesktopIcons\NewStartPanel" -Name "{F02C1A0D-BE21-4350-88B0-7367FC96EF3C}" -value 0
    Set-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\HideDesktopIcons\NewStartPanel" -Name "{5399E694-6CE5-4D6C-8FCE-1D8870FDCBA0}" -value 0
#
# turn off UAC
    Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\policies\system" -Name "ConsentPromptBehaviorAdmin" -Value 0
    Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\policies\system" -Name "PromptOnSecureDesktop" -Value 0

# set system properties/visual effects/adjust for best performance
    New-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\VisualEffects" -Name "VisualFXSetting" -Value 2 -PropertyType "DWORD"
#
# restart windows explorer
    taskkill /F /IM explorer.exe
    explorer.exe
#
# page file
echo "set page file size in next 2 commands"
    if ($ram -le 32768)
	   {
		    $pagefile = Get-WmiObject Win32_ComputerSystem -EnableAllPrivileges
			$pagefile.AutomaticManagedPagefile = $false
			$pagefile.put() | Out-Null
			$pagefileset = Get-WmiObject Win32_pagefilesetting
			$pagefileset.InitialSize = ($ram * 1.5)
			$pagefileset.MaximumSize = ($ram * 1.5)
			$pagefileset.Put() | Out-Null
			write-host "Ram is less than or equal to 32 GB, setting page file to 150% of RAM"
		}
    elseif ($ram -gt 32768)
	   {
		    $pagefile = Get-WmiObject Win32_ComputerSystem -EnableAllPrivileges
			$pagefile.AutomaticManagedPagefile = $false
			$pagefile.put() | Out-Null
			$pagefileset = Get-WmiObject Win32_pagefilesetting
			$pagefileset.InitialSize = (61440)
			$pagefileset.MaximumSize = (61440)
			$pagefileset.Put() | Out-Null
            write-host "Ram is greater than 32 GB, setting page file to 60 GB"
		}
	else
	{
		 write-host "No Change necessary"	
	}
# 
# set time zone and military time display
    Set-TimeZone -Id "$tz Standard Time"
    Set-ItemProperty -Path "HKCU:\Control Panel\International" -Name sShortTime -Value "HH:mm";
echo  "Don't Forget to install VNC and notepad++"
#
# install windows features and open group policy editor and system propties
	Install-WindowsFeature -Name Multipath-IO,Windows-Server-Backup,Telnet-Client
# open system propties, performance options
    SystemPropertiesPerformance.exe
# open desktop icon settings control panel 
    cmd.exe /c "desk.cpl 5,"
#and open local group policy editor 
    gpedit.msc
echo "All Done !! "
echo "reboot the server after all settings are complete !!!"
echo "opening windows explorer to location of extra software"
 $extra = (read-host "where are the extra software programs (such as notepad++) located?")
 explorer $extra
<#
	2019 settings
	CPU reservations are number of logical CPUs x clock speed (2.5 GHz minimum)	
	=====================================================================================================================
	reference doc --  AF7908_WINDOWS_OS_Installation_Guide.pdf	
	=====================================================================================================================
	sconfig	(follow doc AF7908 page 11 section 2.3.2 steps 4-6c)
		2	set hostname if necessary, say no to restart
		4	3	yes	OK	(enable ping)
		4			
		5	A	ok	(win updates)
		7	E	2	(enable RDP)
		8	1		(set IP and DNS if needed)
		4
		9			(set date time timezone)
		6	A		(download and install win updates)
		13	yes		(reboot)
	sconfig
		5	M	OK	(set win updates to Manual)
		15
#>
# Apply the equivalent of the gpedit.msc settings below so this can run unattended.
# 1) Turn off Autoplay for all drives
Try {
	New-Item -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies' -Name 'Explorer' -ErrorAction SilentlyContinue | Out-Null
	Set-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer' -Name 'NoDriveTypeAutoRun' -Value 0xFF -Type DWord -Force
	New-Item -Path 'HKLM:\SOFTWARE\Policies\Microsoft\Windows' -Name 'Explorer' -Force | Out-Null
	Set-ItemProperty -Path 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\Explorer' -Name 'NoDriveTypeAutoRun' -Value 0xFF -Type DWord -Force
	Write-Host "Autoplay disabled (NoDriveTypeAutoRun=0xFF)"
} Catch {
	Write-Warning "Failed to set Autoplay registry keys: $_"
}

# 2) Password policy: Maximum password age = 0 (never expire) and disable complexity.
# Use secedit to apply local security policy changes.
Try {
	$secInf = Join-Path $env:TEMP 'LocalPasswordPolicy.inf'
	@" 
[Unicode]
Unicode=yes
[System Access]
MaximumPasswordAge = 0
PasswordComplexity = 0
"@ | Out-File -FilePath $secInf -Encoding Unicode

	# Apply security policy changes (this will update the local security database)
	secedit.exe /configure /db "$env:windir\security\local.sdb" /cfg $secInf /areas SECURITYPOLICY | Out-Null
	Remove-Item -Path $secInf -Force -ErrorAction SilentlyContinue
	Write-Host "Password policy updated: MaximumPasswordAge=0, PasswordComplexity=0"
} Catch {
	Write-Warning "Failed to apply password policy via secedit: $_"
}

# 3) Network List Manager Policies: set Unidentified Networks to Private via policy registry
Try {
	New-Item -Path 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\NetworkList' -Name 'UnidentifiedNetworks' -Force | Out-Null
	# Use LocationType = 1 for Private (0 = Not configured, 1 = Private, 2 = Public)
	Set-ItemProperty -Path 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\NetworkList\UnidentifiedNetworks' -Name 'LocationType' -Value 1 -Type DWord -Force
	Write-Host "Unidentified Networks policy set to Private (LocationType=1)"
} Catch {
	Write-Warning "Failed to set Unidentified Networks policy: $_"
}

Write-Host "Note: Some policy changes require a reboot or running 'gpupdate /force' to take effect."