Function Set-ScriptVariable {

    ### Info Block ###
    $Script:ErrorActionPreference = "Stop"
    $Script:scriptName = "update-pcoip-agent.ps1"
    $Script:scriptVersion = "2024.02.28"
    Set-StrictMode -Version '2.0'

    Try {

        ### Module import ###
        Import-Module -Name "$PSScriptRoot\Module\WorkspaceScriptModule\WorkspaceScriptModule.psd1"
        $Script:WspApplicable = Test-WspApplicabilityStatus
        Set-ProtocolInstallType -InstallType "$($Script:WspApplicable.InstallType)"

    }
    Catch [System.IO.FileNotFoundException] {

        . "$PSScriptRoot\shared-utils.ps1"
        $Script:WspApplicable = $False

    }

    #Using the new function to determine the workspace type.
    $Script:WorkSpaceType = Get-WorkspaceType
    switch ($Script:WorkSpaceType) {
        "Graphics" { $Script:application = "PCoIPAgent_v2" }
        "Standard" { $Script:application = "PCoIP_agent_installer_cmd" }
    }
    $Script:osType = $(Get-WsOsInfo).OsCaption
    $Script:win2008R2_Or_Win7 = 6
    $Script:packageFileName = "${application}.zip"
    $Script:remoteFileRootPath = "updates/apps/${application}"
    Set-DownloadDirectory
    $Script:downloadDirectory = "$env:ProgramFiles\Amazon\WorkspacesConfig\Temp"
    $Script:EnvVariableName = 'TMP'
    $Script:localPackageFile = ('{0}\{1}' -f $downloadDirectory, $packageFileName)
    #Teradici Log Cleanup configuration
    $Script:teradiciLogFolder = "C:\ProgramData\Teradici\logs"
    $Script:daysToCleanUp = 7

    #BreakingChange N & N-1 versions for Skylight & PCoIP
    $Script:SkylightNversion = "2.6.195.0"
    $Script:PCoIPNversion = "20.10.4"
    $Script:PCoIPN_1_version = "2.7.9.12212"
    # Windows Server 2022 and greater will only support PCoIP Agent 22.04.1 and greater.  Set minimum version to 22.04.1
    $Script:Server202xMinimumVersion = "22.04.1"
    $Script:PCoIPAgentUnSupportedOSRegex = ".*Server 202.*"
    #LogOn Banner enabled Maximum supported versions
    $Script:LogonBannerMaxVersion = "20.10.8"
    $Script:LogonBannerServer202xMaxVersion = "22.04.2"
    #UsbRedirectionKeys
    $Script:UsbRedirectionKeyPath = "HKLM:\SOFTWARE\Policies\Teradici\PCoIP\pcoip_admin"
    $Script:UsbRedirectionKeyPath_NonAdm = "HKLM:\SOFTWARE\Policies\Teradici\PCoIP\pcoip_admin_defaults"
    $Script:UsbRedirectionKeyName = "pcoip.enable_usb"
    $Script:UsbAuthTableKeyName = "pcoip.usb_auth_table"
    $Script:UsbAdmKeyName = "UsbAdm"
    $Script:BackupPath = "HKLM:\SOFTWARE\Amazon\WorkSpacesConfig\$scriptName"
    #Reboot control
    #If upgrade to new ver fails, reboot and try again
    #If upgrade still fails after reboot, install the version before reboot
    $Script:RevertVersionKeyname = "RevertVersion"
    $Script:UpgradeVersionKeyName = "UpgradeVersion"
    $Script:RebootCountKeyName = "RebootCount"
    $Script:RevertBackToOldVer = "RevertBackToOldVer"
    $Script:PCoIPAgentStatusKeyName = "PCoIPAgentStatus" #Values == InstallStarted / UnInstallStarted / UnInstallSucceeded / UnInstallFailed / UpgradeSucceeded / UpgradeFailed / RollbackSuceeded / RollbackFailed
    $Script:PCoIPRebootCountAppName = "PCoIP_RebootCount"
    $Script:rebootWorkSpaceInProgress = 'False'
    # Seamless Migration Variables
    $Script:ConfigFileName = "workspace_config"
    $Script:WorkspaceConfigFileName = "$ConfigFileName.json"
    $Script:LocalTempWorkspaceConfigFilePath = "$Script:downloadDirectory\$WorkspaceConfigFileName"
    $Script:WorkspaceConfigRemoteFileRootPath = "updates/config/$ConfigFileName"
    $script:TaskName = "TeraRestoreWDDMDriver"
    #Disable USB Webcam Redirection
    $Script:DisablePCoIPWebcamRegKeys = @($($UsbRedirectionKeyPath), $($UsbRedirectionKeyPath_NonAdm))
    #Check PCoIP Firewall rules, if doesn't exit, then create
    $Script:FirewallAgentName = "PCoIP - Agent Service"
    $Script:FirewallServerName = "PCoIP - Server"
    $Script:AgentPath = "C:\Program Files\Teradici\PCoIP Agent\bin\pcoip_agent.exe"
    $Script:ServerPath = "C:\Program Files\Teradici\PCoIP Agent\bin\pcoip_server.exe"
    #Adding WSP Manifest Allowed States
    $Script:AllowedState = @("NULL", "StxhdDisabledWspInstalled", "StxhdRemovedWspInstalled")
    # If Get-InstalledWspVersion or Get-InstalledPCoIPVersion returns NULL, the output cannot be converted to [System.Version]
    try {
        [System.Version]$Script:ExistingPCoIPVersion = Get-InstalledPCoIPVersion

    }
    catch {
        Log-Exception $_
        Log-Info "Unable to format existing PCoIP Version to [System.Version]. Retrying to check if existing version is NULL"
        $Script:ExistingPCoIPVersion = Get-InstalledPCoIPVersion
        Log-Info "Existing PCOIP version is : $Script:ExistingPCoIPVersion"
    }

    try {
        [System.Version]$Script:ExistingWSPVersion = Get-InstalledWspVersion

    }
    catch {
        Log-Exception $_
        Log-Info "Unable to format existing WSP Version to [System.Version]. Retrying to check if existing version is NULL"
        $Script:ExistingWSPVersion = Get-InstalledWspVersion
        Log-Info "Existing WSP version is : $Script:ExistingWSPVersion"
    }
}

Function Get-RebootKeyValue {
    return (Get-RegistryKeyValue -KeyPath $BackupPath -KeyName $RebootCountKeyName)
}

Function Get-RevertKeyValue {
    return (Get-RegistryKeyValue -KeyPath $BackupPath -KeyName $RevertBackToOldVer)
}

Function Get-ManifestFile {
    <#
.SYNOPSIS
Downloads a manifest file from the update server as specified
.DESCRIPTION
1. Application = tera_dev_con, downloads the teraDevCon manifest file
2. RevertKey is != 1 OR (RevertKey != Null && InstalledAppVersion != Null), the script will pull down whatever manifest that
   has been specified respecting any overrides found for the machine.
3. RevertKey is = 1 OR (CurrentRebootCountValue = 2 && RevertKey = Null && InstalledAppVersion = Null), and application != tera_dev_con, the script will try
   and find a revertVersion that it can go to. It will first check RevertVersion key for a value, if that is not present
   it will check the rootManifestFile for a RevertVersion xml tag and extract the value, and if that is also not found, it
   will fall back to the hardcoded version of the variable $PCoIPN_1_version
#>
    $RevertKeyValue = Get-RevertKeyValue
    If (($RevertKeyValue -eq 1 -or ($CurrentRebootCountValue -eq 2 -and $Null -eq $RevertKeyValue -and 'NULL' -eq $ExistingVersion)) -and $application -ne "tera_dev_con") {
        Log-Info -Message "RevertKeyVaue = '$RevertKeyValue' and application = '$application' and existing version = '$ExistingVersion'"
        Set-RegistryKeyValue -KeyPath $BackupPath -KeyName $RevertBackToOldVer -PropertyType 'DWord' -Value 1 -Exit
        $Script:RevertVersion = Get-RegistryKeyValue -KeyPath $BackupPath -KeyName $RevertVersionKeyname
        If ($Null -ne $RevertVersion) {
            Log-Info -Message "Revert Version in registry key at location '$BackupPath' is $RevertVersion"
            $remoteFileRootPath = $remoteFileRootPath + "/$RevertVersion"
        }
        Else {
            $success = Get-FileFromRootLocation -Application $Application -ManifestFileName $manifestFileName -LocalFilePath $localManifestFile
            If ($success) {
                Log-Info -Message "Getting revert version from root manifest file"
                $script:RevertVersionInDownloadedManifestFile = Get-RevertVersionFromManifestFile $localManifestFile
                Log-Info -Message "Revert version in root manifest file is '$script:RevertVersionInDownloadedManifestFile'"
            }
            If ($Null -eq $RevertVersionInDownloadedManifestFile) {
                Log-Info -Message "Setting revert version to '$PCoIPN_1_version' using hardcoded version in the script. "
                $remoteFileRootPath = $remoteFileRootPath + "/$PCoIPN_1_version"
            }
            Else {
                $remoteFileRootPath = $remoteFileRootPath + "/$RevertVersionInDownloadedManifestFile"
            }
        }
    }
    Log-Info "Searching for file '$manifestFileName' at location '$remoteFileRootPath' to be downloaded to '$localManifestFile'"
    Download-FileFromUpdateServer -ServerFileRootPath $remoteFileRootPath -FileName $manifestFileName -LocalFilePath $localManifestFile
}

Function Is-SupportedWindowsVersion {
    $osVersion = [Environment]::OSVersion.Version
    $osVersionMajor = $osVersion.Major
    $osVersionMinor = $osVersion.Minor
    Log-Info ("Windows OS Version: major: {0} minor: {1} caption: {2}" -f $osVersionMajor, $osVersionMinor, $Script:osType)
    if (($osVersionMajor -eq 6 -and $osVersionMinor -eq 1) -or ($osVersionMajor -eq 10 -and $osVersionMinor -eq 0 ) -and (-not($Script:osType -like "*Windows 11*"))) {
        return $true, $osVersionMajor, $Script:osType
    }
    else {
        return $false, $osVersionMajor, $Script:osType
    }
}

Function Can-UpdatePCoIPAgent {
    $configurationStatus = Get-SkyLightWorkspaceConfigServiceStatus
    Log-Info ("WorkSpace Configuration state: {0}" -f $configurationStatus)
    $serviceState = Get-Service SkyLightWorkspaceConfigService
    Log-Info ("SkyLightWorkspaceConfigService service state: {0}" -f $serviceState.Status)
    If ($serviceState.Status -ne "Stopped" -and ($configurationStatus -eq 0 -or $configurationStatus -eq 1)) {
        return $false
    }
    else {
        return $true
    }
}

Function Uninstall-PCoIPAgent {
    Log-Info ("Uninstalling PCoIP $ExistingVersion")
    Log-Info ("Start executing '{0}' with parameters '{1}'." -f "$oldUninstallerFile", "/S /NoPostReboot _?=C:\Program Files (x86)\Teradici\PCoIP Agent")
    Update-PCoIP_Update_state -PCoIPUpdate "UnInstallStarted"
    Set-EnvironmentVariableForScope -Name $EnvVariableName -Value $DownloadDirectory -ScriptName $scriptName -Scope Process
    $ExitCode = (Start-Process -FilePath $oldUninstallerFile -ArgumentList "/S /NoPostReboot '_?=C:\Program Files (x86)\Teradici\PCoIP Agent'" -Wait -PassThru).ExitCode
    switch ($ExitCode) {
        0 { $success = "True"; $result = "success"; $script:requireReboot = "False" }
        1 { $success = "False"; $result = "installation aborted by user (user cancel)"; $script:requireReboot = "False" }
        2 { $success = "False"; $result = "installation aborted due to error"; $script:requireReboot = "False" }
        1641 { $success = "True"; $result = "success, but reboot required"; $script:requireReboot = "True" }
        default { $success = "False"; $result = "Unknown ExitCode returned"; $script:requireReboot = "False" }
    }
    Log-Info -Message "Exit code = '$ExitCode'"
    if ("False" -eq $success) {
        Log-Error "Error Uninstalling PCoIP agent $ExistingVersion ExitCode: $ExitCode - $result"
        Update-PCoIP_Update_state -PCoIPUpdate "UnInstallFailed"
    }
    Else {
        Log-Info "Uninstalling PCoIP agent $ExistingVersion ExitCode: $ExitCode - $result"
        Update-PCoIP_Update_state -PCoIPUpdate "UnInstallSucceeded"
    }
    Log-Info ("Exit code: $result")
    return $success, $script:requireReboot
}

Function Install-PCoIPAgent {
<#
Installation of PCoIP & script initiated rollbacks
1. Version comparison between installed PCoIP agent version and version from downloaded manifest file of the application to decide
   if an upgrade/downgrade is needed.
2. If both versions are same, no changes to the installed pcoip agent is made.
3. If it has to go from version A to version B, installation of version B is attempted twice followed by a rollback to version A.
   If the rollback to version B fails, the machine is left will be in a state, where login via PCoIP agent will not work.
4. Example of rollback to version A when version B installation fails
4.1. Attempt 1 - to go from version A to B. set RebootCount = 0 | set RevertVersion = version A
4.2. Attempt 1 fails. set RebootCount = 1. Machine is rebooted.
4.3. Attempt 2 - to go from version A to B. RebootCount = 1
4.4. Attempt 2 fails. set RebootCount = 2, set RevertBackToOldVer = 1. Machine is rebooted.
4.5. Rollback attempt
4.6. If RevertBackToOldVer = 1, check RevertVersion. If RevertVersion is not empty use that value to attempt rollback. If RevertVersion is
     empty, download manifest file of the application from the global root and extract the value from RevertVersion xml tag. If that is also
     empty, use the hardcoded version from $PCoIPN_1_version variable.
4.7. If rollback fails set RebootCount = 3, throw error and exit.
#>
    # Cleaning up Teradici artifacts before installing new PCoIP agent. This is not valid during upgrades to 20.10.4 or higher
    if ($null -eq (Get-WmiObject -Class Win32_Service -Filter "Name='PCoIPAgent'") -and ($VersionInDownloadedManifestFile.Split('.')[0] -lt 20) -and ($ExistingVersion -lt '20.10.4' -or 'NULL' -eq $ExistingVersion)) {
        Clean-Up-Teradici | Out-Null
    }
    Update-PCoIP_Update_state -PCoIPUpdate "InstallStarted"
    Log-Info ("Installing PCoIP '$VersionInDownloadedManifestFile'")
    Set-EnvironmentVariableForScope -Name $EnvVariableName -Value $DownloadDirectory -ScriptName $scriptName -Scope Process
    $ExitCode = (Start-Process -FilePath "${downloadDirectory}\${application}.exe" -ArgumentList "/S /NoPostReboot" -Wait -PassThru).ExitCode
    switch ($ExitCode) {
        0 { $success = "True"; $result = "success"; $script:requireReboot = "False" }
        1 { $success = "False"; $result = "installation aborted by user (user cancel)"; $script:requireReboot = "False" }
        2 { $success = "False"; $result = "installation aborted due to error"; $script:requireReboot = "False" }
        1641 { $success = "True"; $result = "success, but reboot required"; $script:requireReboot = "True" }
        default { $success = "False"; $result = "Unknown ExitCode returned"; $script:requireReboot = "False" }
    }
    Log-Info -Message "Exit code = '$ExitCode'"
    $Script:ExistingVersion = Get-InstalledPCoIPVersion
    If ("True" -eq $success -and 'NULL' -eq $ExistingVersion) {
        $success = "False"
        Log-Info -Message "Exit code was '$ExitCode' but installed PCoIP ver was NULL, setting 'success' to '$success'"
    }
    $CurrentRebootCountValue = Get-RebootKeyValue
    if ("False" -eq $success) {
        Log-Error "Error Installing PCoIP agent '$VersionInDownloadedManifestFile' ExitCode: $ExitCode - $result"
        If (($Null -eq $CurrentRebootCountValue -or $CurrentRebootCountValue -eq 0)) {
            Set-RegistryKeyValue -KeyPath $BackupPath -KeyName $RebootCountKeyName -PropertyType 'DWord' -Value 1 -Exit
            Log-Info -Message "Value of '$RebootCountKeyName' at location '$BackupPath' is set to '$(Get-RebootKeyValue)'"
            Update-RegisteredSkyLightApplicationVersion -ApplicationName $PCoIPRebootCountAppName -Version "1" | Out-Null
            Update-PCoIP_Update_state -PCoIPUpdate "UpgradeFailed"
            $rebootWorkSpaceInProgress = 'True'
            Restart-WKSNow -AppName "PCoIPRebootStatus"
        }
        ElseIf ($CurrentRebootCountValue -eq 1) {
            Set-RegistryKeyValue -KeyPath $BackupPath -KeyName $RebootCountKeyName -PropertyType 'DWord' -Value 2 -Exit
            Set-RegistryKeyValue -KeyPath $BackupPath -KeyName $RevertBackToOldVer -PropertyType 'DWord' -Value 1 -Exit
            Update-RegisteredSkyLightApplicationVersion -ApplicationName $PCoIPRebootCountAppName -Version "2" | Out-Null
            Update-PCoIP_Update_state -PCoIPUpdate "UpgradeFailed"
            $rebootWorkSpaceInProgress = 'True'
            Restart-WKSNow -AppName "PCoIPRebootStatus"
        }
        Else {
            $UpgradeVersion = Get-RegistryKeyValue -KeyPath $BackupPath -KeyName $UpgradeVersionKeyName
            $RevertVersion = Get-RegistryKeyValue -KeyPath $BackupPath -KeyName $RevertVersionKeyname
            Set-RegistryKeyValue -KeyPath $BackupPath -KeyName $RebootCountKeyName -PropertyType 'DWord' -Value 3 -Exit
            Update-RegisteredSkyLightApplicationVersion -ApplicationName $PCoIPRebootCountAppName -Version "3" | Out-Null
            Update-PCoIP_Update_state -PCoIPUpdate "RollbackFailed"
            Throw "Installation of '$UpgradeVersion' failed and also reverting back to '$RevertVersion' failed"
        }
    }
    If (((Get-RevertKeyValue) -eq 1) -and ($VersionInDownloadedManifestFile -eq $ExistingVersion)) {
        Update-PCoIP_Update_state -PCoIPUpdate "RollbackSucceeded"
        Update-RegisteredSkyLightApplicationVersion -ApplicationName "RollbackSucceeded" -Version "$Script:ExistingVersion" | Out-Null
    }
    ElseIf ($VersionInDownloadedManifestFile -eq $ExistingVersion) {
        Update-PCoIP_Update_state -PCoIPUpdate "UpgradeSucceeded"
        Update-RegisteredSkyLightApplicationVersion -ApplicationName "UpgradeSucceeded" -Version "$Script:ExistingVersion" | Out-Null
        try {
            Log-Info "Removing DetectIdleShutdownAgentInstall from HKLM:\Software\Teradici\PCoIPAgent\DetectIdleShutdownAgentInstall."
            Remove-Item "HKLM:\Software\Teradici\PCoIPAgent\DetectIdleShutdownAgentInstall" -Force
            Log-Info "Removed DetectIdleShutdownAgentInstall from HKLM:\Software\Teradici\PCoIPAgent\DetectIdleShutdownAgentInstall. As this is not required in 22.04.1 or higher or lower"
        }
        catch {
            Log-Exception $_
        }
    }
    Else {
        Update-PCoIP_Update_state -PCoIPUpdate "Unknown"
    }
    Log-Info ("Exit code: $result")
    return $success, $script:requireReboot
}

Function Backup-PCoIPPrintDefault {
    $BackupPath = "C:\Program Files\Amazon\WorkSpacesConfig\Backup"
    $Config_Path = "$BackupPath\$($existingVersion.Replace('.','_')).reg"
    $Path = "HKLM:\SOFTWARE\Policies\Teradici\PCoIP\pcoip_admin_defaults"
    $RegPath = $Path.Replace(":", "")
    $Name = "pcoip.remote_printing_enabled"
    $Remote_Printing = Get-RegistryKeyValue -KeyPath $Path -Keyname $Name
    if ($existingVersion -match "^2.1\..*") {
        if (-not(Test-Path $BackupPath)) {
            New-Item -Path $BackupPath -ItemType Directory
        }
        Log-Info ("Backing up $Path registry")
        Invoke-Expression "cmd.exe /c reg export $RegPath '$Config_Path' /y"
    }
    if ($existingVersion -eq "2.1.1.1270" -and ($Remote_Printing -eq 1 -or  $Null -eq $Remote_Printing )) {
        if (-not($Remote_Printing)) {
            $Remote_Printing = "Null"
        }
        Set-RegistryKeyValue -KeyPath $Path -KeyName $Name -PropertyType 'DWORD' -Value 2
        Set-RegistryKeyValue -KeyPath $Path -KeyName "pcoip.enable_default_printer" -PropertyType 'DWORD' -Value 1
    }

    if ($existingVersion -eq "2.1.1.1505" -and $Remote_Printing -eq 1) {
        Log-Info ("RegKey $name is set to $Remote_Printing")
        Log-Info ("Setting $Path\$Name to 2")
        Set-ItemProperty -Path $Path -Name $Name -Value 2
    }

}

Function Restore-PCoIPPrintBackup {
    $Print_Reg_Backup = "C:\Program Files\Amazon\WorkSpacesConfig\Backup\$($VersionInDownloadedManifestFileWithoutBuildNumber.Replace('.','_')).reg"
    if (Test-Path $Print_Reg_Backup) {
        Log-Info ("Detected PCoIP printing registry backup. Importing the configuration")
        Invoke-Expression "cmd.exe /c reg delete $RegPath /f /va
        regedit /s '$Print_Reg_Backup'"
    }
}

Function Has-ProvisioningCompleted {
    $configurationStatus = Get-SkyLightWorkspaceConfigServiceStatus
    Log-Info ("WorkSpace Configuration state: {0}" -f $configurationStatus)
    If ($configurationStatus -eq 2) {
        return $true
    }
    else {
        return $false
    }
}

Function Check-EnoughDiskSpaceAvailable {
    $freeSpaceC = Get-WmiObject -Query "select * from Win32_LogicalDisk where DeviceID = 'C:'"
    $freeSpaceC = $freeSpaceC.FreeSpace / 1GB
    $freeSpaceD = Get-WmiObject -Query "select * from Win32_LogicalDisk where DeviceID = 'D:'"
    $freeSpaceD = $freeSpaceD.FreeSpace / 1GB
    Log-Info ("Available Disk Space on C: drive: {0} GBs" -f $freeSpaceC)
    Log-Info ("Available Disk Space on D: drive: {0} GBs" -f $freeSpaceD)
    if ($freeSpaceC -ge 2) {
        return $true
    }
    else {
        return $false
    }
}

Function Extract-PcoipPackage {
    param([String]$localPackageFile, [String] $downloadDirectory)

    #Extracting the downloaded PCoIP zip file before uninstalling the exisiting agent
    Log-Info ("Extract package '{0}' into '{1}' before uninstalling current PCoIP agent." -f $localPackageFile, $downloadDirectory)
    $initializationScript = { Import-Module -Name "C:\Program Files\Amazon\WorkspacesConfig\Scripts\Module\WorkspaceScriptModule\WorkspaceScriptModule.psd1" }
    $scriptBlock = { param($arg1, $arg2) Extract-ZipV2 -SourceZipFile $arg1 -DestinationFolder $arg2 }
    $extractJob = Start-Job -InitializationScript $initializationScript -ScriptBlock $scriptBlock -ArgumentList ($localPackageFile, $downloadDirectory)

    # The logic here is wait for a period of time(set by timeout) and then proceed to next line, while
    # the job still running on the background
    $extractJob | Wait-Job -Timeout 10

    # If the job compeleted normally, return
    if ($extractJob.State -eq 'Completed') {
        Log-Info "Extraction job completed"
        return
    }
    # If the job is still running, manually stop the job and handle it as a timeout
    elseif ($extractJob.State -eq 'Running') {
        $extractJob.StopJob()
        Log-Error "Extraction job timed-out"
    }
    # If the job did not complete for any other reason, log the info
    else {
        Log-Error ("Extraction job incomplete. Job StateInfo: '{0}', Reason:'{1}'" -f $extractJob.JobStateInfo, $extractJob.ChildJobs[0].JobStateInfo.Reason)
    }
    # Update the .application file for tracking
    Update-RegisteredSkyLightApplicationVersion -ApplicationName "PCoIP_Agent_Update_Failure" -Version "Failure-Extract-Zip" | Out-Null
    # Throw an Exception to make sure the script won't start uninstalling existing PCoIP agent with an incomplete extraction of new PCoIP agent pacakge
    throw ("Cannot extract PCoIP package '{0}' into '{1}' before uninstalling current PCoIP agent." -f $localPackageFile, $downloadDirectory)
}

Function Configure-PCoIPAgent {
    Log-Info ("Start to configure '{0}' of version '{1}'" -f $application, $ExistingVersion)
    # Verifying the PCoIP Agent Version that matches from RT10 to RT22
    if ($ExistingVersion -match "^(2\.|20\.|22\.)[\d+.]+") {
        try {
            $OSVerionObj = Is-SupportedWindowsVersion
            if(($OSVerionObj[1] -eq $Script:win2008R2_Or_Win7)){
                Log-Info "Enabling Software Secure Attention Sequence"
                Set-ItemProperty -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Policies\System" -Name "SoftwareSASGeneration" -Value 1
                Update-RegisteredSkyLightApplicationVersion -ApplicationName "SoftwareSASGeneration" -Version "True" | Out-Null
            }
            else{
                Log-Info "Removing DisableCAD Registry Key"
                Remove-RegistryKeyValue -KeyPath "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System" -KeyName "DisableCAD"
                Update-RegisteredSkyLightApplicationVersion -ApplicationName "DisableCAD" -Version "False" | Out-Null
                Log-Info "Removing SoftwareSASGeneration Registry Key"
                Remove-RegistryKeyValue -KeyPath "HKLM:\Software\Microsoft\Windows\CurrentVersion\Policies\System" -KeyName "SoftwareSASGeneration"
                Update-RegisteredSkyLightApplicationVersion -ApplicationName "SoftwareSASGeneration" -Version "False" | Out-Null
                Log-Info "Removing HideFastUserSwitching Registry Key"
                Remove-RegistryKeyValue -KeyPath "HKLM:\Software\Microsoft\Windows\CurrentVersion\Policies\System" -KeyName "HideFastUserSwitching"
                Update-RegisteredSkyLightApplicationVersion -ApplicationName "HideFastUserSwitching" -Version "False" | Out-Null
            }
            Log-Info "Reset Power Plan scheme."
            &powercfg.exe -restoredefaultschemes
            if ($LASTEXITCODE -ne 0) {
                Log-Error ("Fail to reset Power Plan scheme. Exit code {0}." -f $LASTEXITCODE)
            }
            else {
                Log-Info ("Successfully reset Power Plan scheme. Exit code {0}." -f $LASTEXITCODE)
            }
            Log-Info "Setting Power Plan to 'High performance' with InstanceID 'Microsoft:PowerPlan\{8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c}'."
            $PowerPlanObj = Get-CimInstance -Name root\cimv2\power -Class win32_PowerPlan -Filter "InstanceID = 'Microsoft:PowerPlan\\{8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c}'" -OperationTimeoutSec 10
            if ($null -ne $PowerPlanObj) {
                $Exitcode = (Start-Process "C:\Windows\System32\powercfg.exe" -ArgumentList "/s 8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c" -Wait -PassThru).ExitCode
                if ($Exitcode -eq 0) {
                    Log-Info "Powerplan Set to high Performance was sucessfully"
                    Update-RegisteredSkyLightApplicationVersion -ApplicationName "Powerplan_HighPerformance" -Version "Success" | Out-Null
                }
                else {
                    Log-Error "Powerplan Set to high Performance failed with exit code $($Exitcode)"
                    Update-RegisteredSkyLightApplicationVersion -ApplicationName "Powerplan_HighPerformance" -Version "Failed" | Out-Null
                }
            }
            else {
                Log-Error "Cannot find 'High performance' power plan with InstanceID as 'Microsoft:PowerPlan\{8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c}'! Skip setting active power plan to 'High performance'."
            }
            Log-Info "Set MaxCachedIcons to 2048"
            Set-ItemProperty -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Explorer" -Name "MaxCachedIcons" -Value "2048"
            Add-PCoIP_Legacy_regkey
        }
        catch {
            Log-Exception $_
        }
    }
    else {
        Log-Info "Unable to find the Version: $ExistingVersion"
    }
    Set-ProcessDumpRegKey
    Log-Info ("Finish configuring '{0}' of version '{1}'" -f $application, $ExistingVersion)
}

Function Update-PCoIPAgent_Legacy {
    Log-info "Using function : Update-PCoIPAgent_Legacy"
    # Stop Skylight Agent
    try {
        Stop-SkyLightWorkspaceConfigService
        # Stop PCoIP agent
        Stop-PCoIPAgentService
        Log-Info ("PCoIP agent service status: $($(Get-Service PCoIPAgent).Status)")
    }
    catch {
        Log-Exception $_
    }
    # Extract new PCoIP package before unistalling old PCoIP agent
    Extract-PcoipPackage -localPackageFile $localPackageFile -downloadDirectory $downloadDirectory
    #Uninstalling PCoIP Agent
    $oldUninstallerDirectory = "C:\Program Files (x86)\Teradici\PCoIP Agent"
    $oldUninstallerFile = "C:\Program Files (x86)\Teradici\PCoIP Agent\uninst.exe"
    # Declaring default PCoIP update variables
    if (Test-Path -Path $oldUninstallerDirectory) {
        if (Test-Path -Path $oldUninstallerFile) {
            try {
                #Backing up remote printing configuration and translate them its equivalents for RT10+
                Backup-PCoIPPrintDefault
                # Uninstalling old PCoIP agent
                $UpdateResult = Uninstall-PCoIPAgent
            }
            catch {
                Log-Error ("Failed to uninstall PCoIP agent")
                Log-Exception $_
            }
            Log-Info ("Finished uninstalling PCoIP $ExistingVersion")
        }
        else {
            Log-Info ("Unable to find $oldUninstallerFile")
            $UpdateResult = @("True","False")
        }
    }
    else {
        Log-Info ("Unable to find $oldUninstallerDirectory")
        $UpdateResult = @("True","False")
    }
    #Installing new PCoIP Agent if the uninstall=success or there is no PCoIP agent present
    if ("True" -eq $UpdateResult[0] -and ($null -eq (Get-WmiObject -Class Win32_Service -Filter "Name='PCoIPAgent'"))) {
        if (Test-Path "${downloadDirectory}\${application}.exe") {
            $credentialsDllFilepath = "C:\Windows\system32\pcoip_credential_provider.dll"
            $renamedcredentialsDllFilepath = "C:\Windows\system32\pcoip_credential_provider.bak"
            if (Test-Path -Path $credentialsDllFilepath) {
                Log-Info ("Starting to rename $credentialsDllFilepath")
                try {
                    if (Test-Path -Path $renamedcredentialsDllFilepath) {
                        Log-Info ("Removing bak Creds File")
                        Remove-Item $renamedcredentialsDllFilepath -Force
                    }
                    else {
                        Log-info "Path $renamedcredentialsDllFilepath is False"
                    }
                    Rename-Item -Path $credentialsDllFilepath -NewName $renamedcredentialsDllFilepath
                    Log-Info ("Finished renaming $credentialsDllFilepath to $renamedcredentialsDllFilepath")
                }
                catch {
                    Log-Error ("Exception in failing to move creds dll")
                    Log-Exception $_
                }
            }
            else {
                Log-Info "Path $credentialsDllFilepath is False"
            }
            try {
                $UpdateResult = "False"
                $UpdateResult = Install-PCoIPAgent
                Log-Info ("Installed version of '{0}' is '{1}'." -f $application, $ExistingVersion)
                if ($VersionInDownloadedManifestFileWithoutBuildNumber -ne $ExistingVersion) {
                    throw ("The version of '{0}' after updating is '{1}', different from expected version '{2}'. The update of '{0}' failed!" -f $application, $ExistingVersion, $VersionInDownloadedManifestFileWithoutBuildNumber)
                }
                else {
                    Restore-PCoIPPrintBackup
                }
                Log-Info ("Update registered version of application '{0}' in .applications file to '{1}'." -f $application, $VersionInDownloadedManifestFile)
                Update-RegisteredSkyLightApplicationVersion -ApplicationName $application -Version "$VersionInDownloadedManifestFile" | Out-Null
                Log-Info ("Finish updating '{0}' to version '{1}'." -f $application, $VersionInDownloadedManifestFileWithoutBuildNumber)
            }
            catch {
                Log-Error ("Failed to update to agent")
                Log-ErrorThenExit -Message $_
            }
            Log-Info ("Finish executing '{0}'." -f "${downloadDirectory}\${application}.exe")
        }
        else {
            throw "Cannot find '${downloadDirectory}\${application}.exe' for updating ${application}!"
        }
        #Getting the reboot tag from the PCoIP installation ExitCode
        $script:requireReboot = $UpdateResult[1]
        $script:didsucceed = $UpdateResult[0]
        Log-Info ("Require Reboot: '{0}'" -f $script:requireReboot)
        $script:provisioningCompleted = Has-ProvisioningCompleted
        Log-Info ("WorkSpace Provisioning Completed: '{0}'" -f $script:provisioningCompleted)
    }
    else {
        Log-Error ("Cannot uninstall PCoIP $existingVersion. Skip updating.")
    }
}

Function Update-PCoIPAgent_Inplace {
    Log-info "Using function : Update-PCoIPAgent_Inplace"
    # Stop Skylight Agent
    try {
        Stop-SkyLightWorkspaceConfigService
    }
    catch {
        Log-Exception $_
    }
    #Installing new PCoIP Agent on top old version
    Log-Info ("Extract package '{0}' of '{1}' into '{2}'." -f $localPackageFile, $application, $downloadDirectory)
    Extract-ZipV2 -SourceZipFile $localPackageFile -DestinationFolder $downloadDirectory
    if (Test-Path "${downloadDirectory}\${application}.exe") {
        try {
            $UpdateResult = Install-PCoIPAgent
            Log-Info ("Installed version of '{0}' is '{1}'." -f $application, $ExistingVersion)
            if ($VersionInDownloadedManifestFileWithoutBuildNumber -ne $ExistingVersion) {
                throw ("The version of '{0}' after updating is '{1}', different from expected version '{2}'. The update of '{0}' failed!" -f $application, $ExistingVersion, $VersionInDownloadedManifestFileWithoutBuildNumber)
            }
            else {
                Restore-PCoIPPrintBackup
            }
            Log-Info ("Update registered version of application '{0}' in .applications file to '{1}'." -f $application, $VersionInDownloadedManifestFile)
            Update-RegisteredSkyLightApplicationVersion -ApplicationName $application -Version "$VersionInDownloadedManifestFile" | Out-Null
            Log-Info ("Finish updating '{0}' to version '{1}'." -f $application, $VersionInDownloadedManifestFileWithoutBuildNumber)
        }
        catch {
            Log-Error ("Failed to update to agent")
            Log-ErrorThenExit -Message $_
        }
        Log-Info ("Finish executing '{0}'." -f "${downloadDirectory}\${application}.exe")
    }
    else {
        throw "Cannot find '${downloadDirectory}\${application}.exe' for updating ${application}!"
    }

    #Getting the reboot tag from the PCoIP installation ExitCode
    $script:requireReboot = $UpdateResult[1]
    $script:didsucceed = $UpdateResult[0]
    Log-Info "Require Reboot: $script:requireReboot"
    $script:provisioningCompleted = Has-ProvisioningCompleted
    Log-Info "WorkSpace Provisioning Completed: $script:provisioningCompleted"
}

Function Remove-TrackingKey {
    #Remove Tracking Keys
    ForEach ($Key in @("$RevertVersionKeyname", "$UpgradeVersionKeyName", "$RebootCountKeyName", "$RevertBackToOldVer")) {
        Remove-RegistryKeyValue -KeyPath "$BackupPath" -KeyName $Key -Exit
    }
    #Clear .application file metrics
    ForEach ($Key in @('PCoIP_RebootCount', 'PCoIP_Agent_Update', 'PCoIP_Agent_Update_Failure')) {
        Log-Info -Message "Removing '$Key' from .applications file"
        Remove-MatchingSkylightApplicationVersion -Application $Key | Out-Null
    }
}

Function Set-PCoIPFirewallRules {
    Try {
        New-NetFirewallRule  -DisplayName $($Script:FirewallAgentName) -Description 'Allows the PCoIP Agent service to receive connections.' -Group 'Teradici PCoIP' -Profile Domain, Private, Public -Enabled True -Action Allow -Program $($Script:AgentPath) -LocalAddress Any -RemoteAddress Any -Protocol TCP -LocalPort 60433, 4172 -RemotePort Any | Out-Null
        Log-Info "Added PCoIP Inbound Firewall Rules for $((Get-NetFirewallRule -DisplayName $Script:FirewallAgentName -ErrorAction SilentlyContinue).DisplayName)."
        New-NetFirewallRule  -DisplayName $($Script:FirewallServerName) -Description 'Allows the PCoIP Server to receive connections from PCoIP Clients.' -Group 'Teradici PCoIP' -Profile Domain, Private, Public -Enabled True -Action Allow -Program $($Script:ServerPath) -LocalAddress Any -RemoteAddress Any -Protocol UDP -LocalPort 4172 -RemotePort Any | Out-Null
        Log-Info "Added PCoIP Inbound Firewall Rules for $((Get-NetFirewallRule -DisplayName $Script:FirewallServerName -ErrorAction SilentlyContinue).DisplayName)."
        Update-RegisteredSkyLightApplicationVersion -ApplicationName "PCOIPFireWallRulesFix" -Version "Success" | Out-Null
    }
    Catch {
        Update-RegisteredSkyLightApplicationVersion -ApplicationName "PCOIPFireWallRulesFix" -Version "Failed" | Out-Null
        Log-Exception $_
    }
}

Function Get-PCoIPFirewallRules {
    Try {
        If ($($Script:ExistingPCoIPVersion.Major) -ge '22') {
            Log-Info "Verifying PCoIP Inbound Firewall Rules Enabled or not"
            $GetAgentFirewall = (Get-NetFirewallRule -DisplayName $Script:FirewallAgentName -ErrorAction SilentlyContinue)
            $GetServerFirewall = (Get-NetFirewallRule -DisplayName $Script:FirewallServerName -ErrorAction SilentlyContinue)
            If (($GetAgentFirewall).Enabled -eq 'True' -and ($GetServerFirewall).Enabled -eq 'True' ) {
                Log-Info "PCoIP Inbound Firewall Exist"
                Update-RegisteredSkyLightApplicationVersion -ApplicationName "PCOIPFireWallRules" -Version "True" | Out-Null
            }
            Else {
                Log-Info "PCoIP Firewall Doesn't exist, Hence creating a new inbound firewall rules"
                Update-RegisteredSkyLightApplicationVersion -ApplicationName "PCOIPFireWallRules" -Version "False" | Out-Null
                Set-PCoIPFirewallRules
            }
        }
        Else {
            Log-Info "PCoIP Agent Version is '$($Script:ExistingPCoIPVersion)', hence not required to set Inbound Firewall rules"
        }
    }
    Catch {
        Update-RegisteredSkyLightApplicationVersion -ApplicationName "PCOIPFireWallRules" -Version "Error" | Out-Null
        Log-Exception $_
    }
}

Function Is-SkyLightVersionSupportedForPCoIPUpgrade {
<#
This function checks for minimum version of skylight agent and will evaluate if PCoIP Agent can be updated.
This is only used for older workspaces (Windows 10) that might have skylight agents less than version "2.6.195.0"
#>
    param([String]$installedSkyLightVersion)
    if ([System.Version]$installedSkyLightVersion -gt [System.Version]$Script:SkylightNversion) {
        return $True
    }
    else {
        return $False
    }
}

Function Is-PCoIPVersionUpgradeSupportedForOS {
<#
This function checks for minimum version of pcoip agent and will evaluate if PCoIP Agent can be updated.
This is only used for older workspaces (Windows 7 and Windows 10) that might have very old skylight agents or are windows 7,2008R2
and have a PCoIP agent that might attempt to upgrade beyond 20.10.4 which isn't supported by that OS.
#>
   param([String]$manifestPCoIPVersion)
   if ([System.Version]$manifestPCoIPVersion -ge [System.Version]$Script:PCoIPNversion) {
        return $True
    }
    else {
        return $False
    }
}

Function Update-PCoIPAgent {

    $Script:CurrentRebootCountValue = Get-RebootKeyValue
    If ($Null -eq $CurrentRebootCountValue) {
        Set-RegistryKeyValue -KeyPath $BackupPath -KeyName $RebootCountKeyName -PropertyType 'DWord' -Value 0 -Exit
    }
    Else {
        Log-Info -Message "Current value of '$RebootCountKeyName' is '$CurrentRebootCountValue'"
    }
    $Script:ExistingVersion = Get-InstalledPCoIPVersion
    Get-PCoIPFirewallRules
    Log-Info ("The current installed version of {0}: {1}" -f $application, $ExistingVersion)
    # Check if customer NIC is present
    $customerInterfaceMac = Get-CustomerInterfaceMac
    if ($null -eq $customerInterfaceMac) {
        Log-Error ("Customer network interface is still not available. Skip updating {0}." -f $application)
        return
    }

    try {
        # Set manifest suffix for WIN7 and WIN10
        switch ($($windowsVersionSupported[1])) {
            6 { $OSsuffix = "" }
            10 { $OSsuffix = "_WIN10" }
        }
        $manifestFileName = "${application}.zip.manifest$OSsuffix.xml"
        $localManifestFile = ('{0}\{1}' -f $downloadDirectory, $manifestFileName)
        Get-ManifestFile
    }
    catch {
        Log-Error ("Cannot download {0} manifest file '{1}' from Update Server.  Skip updating {0}." -f $application, $manifestFileName)
        return
    }

    # New-Old Manifest file check
    $VersionInDownloadedManifestFile = Get-VersionFromManifestFile $localManifestFile
    # This has to stay as String or else the final install check will fail because leading 0 in version will be stripped
    $VersionInDownloadedManifestFileWithoutBuildNumber = ([String]$VersionInDownloadedManifestFile -replace '_.*', '')
    Log-Info ("Version in downloaded manifest file at '{0}': '{1}'" -f $localManifestFile, $VersionInDownloadedManifestFile)
    Log-Info ("Version in downloaded manifest file at '{0}' (no build number): {1}" -f $localManifestFile, $VersionInDownloadedManifestFileWithoutBuildNumber)
    $StringManifestVersion = $VersionInDownloadedManifestFile.Split(".")
    #reset the pcoip version that needs to be installed on the Workspace that does not have PCoIP agent installed
    if(('NULL' -eq $ExistingVersion) -and ($Script:LogOnBannerExists -eq $true) -and ($StringManifestVersion[0] -gt 20)){
        Log-Info -Message "Current version of PCoIP agent installed on the box is 'NULL'. Log on Banner is enabled on the WorkSpace. Verifying agent comptability."
        If($Script:osType -match $Script:PCoIPAgentUnSupportedOSRegex){
            $LogOnBannerSupportedVersion = $Script:LogonBannerServer202xMaxVersion
        }else{
            $LogOnBannerSupportedVersion = $Script:LogonBannerMaxVersion
        }
        Log-Info -Message "Compatible PCoIP agent for '$Script:osType' WorkSpace with LogOn banner enabled is '$LogOnBannerSupportedVersion'."
        If (Test-Path -Path $localManifestFile) {
            Remove-Item -Path $localManifestFile -Force -ErrorAction SilentlyContinue
        }
        Download-FileFromUpdateServer -ServerFileRootPath "${remoteFileRootPath}/$LogOnBannerSupportedVersion" -FileName $manifestFileName -LocalFilePath $localManifestFile
        $VersionInDownloadedManifestFile = Get-VersionFromManifestFile $localManifestFile
        $VersionInDownloadedManifestFileWithoutBuildNumber = ([String]$VersionInDownloadedManifestFile -replace '_.*', '')
        $StringManifestVersion = $VersionInDownloadedManifestFile.Split(".")
    }
    $IntManifestVersion  = [int]$StringManifestVersion[0] * 100 + [int]$StringManifestVersion[1]
    $InstalledSkyLightVersion = Get-RegisteredSkyLightWorkspaceConfigServiceVersion
    $RequiredMinSkyLightVersion = Is-SkyLightVersionSupportedForPCoIPUpgrade $InstalledSkyLightVersion
    $RequiredMinPCoIPVersion = Is-PCoIPVersionUpgradeSupportedForOS $VersionInDownloadedManifestFile
    # Check for required minimum version of skylight agent and pcoip agent for Windows 10
    If ( -not $RequiredMinSkyLightVersion -and $RequiredMinPCoIPVersion -and $windowsVersionSupported[1] -eq 10) {
        Log-Info -Message "OS version detected is '$($windowsVersionSupported[1])'.Version of PCoIP in downloaded manifest file is '$VersionInDownloadedManifestFile' and installed version of Skylight is '$InstalledSkyLightVersion'. Required minimum version of Skylight for upgrading/installing PCoIP '$VersionInDownloadedManifestFile' is '$SkylightNversion'. Exiting out of PCoIP Agent Update Check"
        return
    }
    # Check for required minimum version of skylight agent and pcoip agent for Windows 7/2008R2
    ElseIf ($windowsVersionSupported[1] -eq 6 -and $RequiredMinPCoIPVersion) {
        Log-Info -Message "Upgrading/Installing PCoIP version '$VersionInDownloadedManifestFile' is not supported at the moment on Win7 like WorkSpaces. Exiting out of PCoIP Agent Update Check"
        return
    }
    Else {
        Log-Info -Message "OS version detected is '$($windowsVersionSupported[1])'.Installed Skylight version is '$InstalledSkyLightVersion' and PCoIP version in manifest file is '$VersionInDownloadedManifestFile'. Proceeding with PCoIP Agent Update Check"
    }

    If ('NULL' -eq $ExistingVersion) {
        Log-Info -Message "Current version of PCoIP agent installed on the box is 'NULL'. Proceeding with PCoIP Agent Update Check to see if update is applicable"
    }
    ElseIf(($Script:LogOnBannerExists -eq $true) -and ($StringManifestVersion[0] -gt 20)){
        Log-Info -Message "LogOn banner enabled on the WorkSpace and version from manifest file is greater than RT20 -'$VersionInDownloadedManifestFile'. Skip updating the agent."
        Update-RegisteredSkyLightApplicationVersion -ApplicationName "SkipPCoIPUpgrade" -Version "LogOnBannerDetected" | Out-Null
        return
    }
    Else {
        if (($VersionInDownloadedManifestFileWithoutBuildNumber -eq $ExistingVersion) -and ($VersionInDownloadedManifestFile -eq $ExistingVersion)) {
            Log-Info ("The existing version of '{0}' is '{1}' and the version in downloaded manifest file is '{2}'. Skip updating '{0}'." -f $application, $ExistingVersion, $VersionInDownloadedManifestFileWithoutBuildNumber)
            Log-Info -Message "Updating .applications file with keyname '$application' and value '$VersionInDownloadedManifestFile'"
            Update-RegisteredSkyLightApplicationVersion -ApplicationName $application -Version "$VersionInDownloadedManifestFile" | Out-Null
            Remove-TrackingKey
            return
        }
    }

    # Config check
    $isEnough = Check-EnoughDiskSpaceAvailable
    if (-not$isEnough -and $VersionInDownloadedManifestFileWithoutBuildNumber -match "^2\...*") {
        Log-Error ("Less than 2GB disk space available on C: drive. No upgrade to Console agent is allowed")
        return
    }

    # Download and application validation
    Log-Info ("Different '{0}' version found. Existing version is '{1}'. The version in downloaded manifest file is '{2}'. Proceeding with further checks for '{0}'." -f $application, $ExistingVersion, $VersionInDownloadedManifestFileWithoutBuildNumber)
    $dsaSignature = Get-SignatureFromManifestFile $localManifestFile
    If (Test-Path -Path $localPackageFile) {
        Remove-Item -Path $localPackageFile -Force -ErrorAction SilentlyContinue
    }
    Download-FileFromUpdateServer -ServerFileRootPath "${remoteFileRootPath}/$VersionInDownloadedManifestFile" -FileName $packageFileName -LocalFilePath $localPackageFile
    if (-not (Test-Path $localPackageFile)) {
        Log-Error ("Downloaded '{0}' package file '{1}' can't be found unexpectedly." -f $application, $localPackageFile)
        return
    }
    Log-Info ("Verify signature of downloaded '{0}' package file '{1}' against WorkSpaces certificate." -f $application, $localPackageFile)
    $workspacesCertThumbprint = Get-WorkSpacesCertThumbprint
    $verified = Verify-PackageFileSignature -PackageFilePath $localPackageFile -Signature $dsaSignature -CertThumbprint $workspacesCertThumbprint
    if (-not $verified) {
        Log-Error ("Downloaded package file '{0}' failed signature validation. Stop updating '{1}'." -f $localPackageFile, $application)
        Remove-Item $localPackageFile -Force -ErrorAction SilentlyContinue
        return
    }

    $canUpdatePCoIPAgent = Can-UpdatePCoIPAgent
    if (-not $canUpdatePCoIPAgent) {
        Log-Info ("Cannot update PCoIP Agent at this moment. Skip updating '{0}'!" -f $application)
        return
    }
    Log-Info "Get current status of service of SkyLightWorkspaceConfigService."
    $previousSkyLightWorkspaceConfigServiceState = (Get-Service SkyLightWorkspaceConfigService).Status
    Log-Info ("Current status of service of SkyLightWorkspaceConfigService is '{0}'." -f $previousSkyLightWorkspaceConfigServiceState)

    try {
        #Remove old logs
        Remove-OldTeradiciLogFile -Path $teradiciLogFolder -OlderThanInDays $daysToCleanUp

        #Backup UsbRedirection Setting
        Backup-UsbRedirectionSetting

        #Backup PCoIPRegKeys Setting
        Backup-PCoIPRegKey

        #Backup the current version and upgrade version numbers
        If (($Null -eq $CurrentRebootCountValue -or 0 -eq $CurrentRebootCountValue) -and ("NULL" -ne $ExistingVersion)) {
            Set-RegistryKeyValue -KeyPath $BackupPath -KeyName $RevertVersionKeyname -PropertyType 'String' -Value "$ExistingVersion" -Exit
            Set-RegistryKeyValue -KeyPath $BackupPath -KeyName $UpgradeVersionKeyName -PropertyType 'String' -Value "$VersionInDownloadedManifestFile" -Exit
        }
        Else {
            Log-Info -Message "Current value of '$RebootCountKeyName' is '$CurrentRebootCountValue'"
        }

        # Spliting the update path for WIN7 and WIN10.  Checking for unsupported versions of PCoIP Agents (less than 22.04.1) for Windows Server 2022 or greater
        Log-Info "Updating OS version $($windowsVersionSupported[1]) to PCoIP version $VersionInDownloadedManifestFile"
        if (([System.Version]$VersionInDownloadedManifestFile -lt [System.Version]$Script:Server202xMinimumVersion) -and ($Script:osType -match $Script:PCoIPAgentUnSupportedOSRegex)) {
            Log-Info -Message "PCoIP Agent '$VersionInDownloadedManifestFile' is not supported for '$Script:osType'.  Skipping downgrade"
        }
        elseif (($windowsVersionSupported[1] -eq "10" -and $IntManifestVersion  -ge "207") -or ($WorkSpaceType -eq "Graphics")) {
            Update-PCoIPAgent_Inplace
        }
        else {
            Update-PCoIPAgent_Legacy
        }
        if ("False" -eq $script:requireReboot -and "True" -eq $script:didsucceed) {
            Set-UsbRedirectionSetting
            Restore-PCoIPRegKey -RegKeyPath $Script:downloadDirectory
            Log-Info -Message "Updating .applications file with keyname '$application' and value '$VersionInDownloadedManifestFile'"
            Update-RegisteredSkyLightApplicationVersion -ApplicationName $application -Version "$VersionInDownloadedManifestFile" | Out-Null
            Remove-TrackingKey
        }
        if ("True" -eq $script:requireReboot -and $script:provisioningCompleted) {
            Configure-PCoIPAgent
            Set-UsbRedirectionSetting
            Restore-PCoIPRegKey -RegKeyPath $Script:downloadDirectory
            Log-Info ("Reboot WorkSpace after installing and configuring '{0}' of version '{1}'." -f $application, $VersionInDownloadedManifestFileWithoutBuildNumber)
            $rebootWorkSpaceInProgress = 'True'
            Restart-WKSNow -AppName "PCoIPRebootStatus"
        }
    }
    catch {
        Log-Exception $_
    }
    finally {
        if (-not ('True' -eq $rebootWorkSpaceInProgress)) {
            try {
                Start-PCoIPAgentService
            }
            catch {
                Log-Exception $_
            }
            if ($previousSkyLightWorkspaceConfigServiceState -ne "Stopped") {
                try {
                    Start-SkyLightWorkspaceConfigService
                }
                catch {
                    Log-Exception $_
                }
            }
        }
    }
}

Function Remove-Tree($removePath) {
    Remove-Item $removePath -Force -Recurse -ErrorAction silentlycontinue
    if (Test-Path "$removePath\" -ErrorAction silentlycontinue) {
        $folders = Get-ChildItem -Path $removePath -Directory -Force
        ForEach ($folder in $folders) {
            Remove-Tree $folder.FullName
        }
        $files = Get-ChildItem -Path $removePath -File -Force
        ForEach ($file in $files) {
            Remove-Item $file.FullName -Force
        }
        if (Test-Path "$removePath\" -ErrorAction silentlycontinue) {
            Remove-Item $removePath -Force -Recurse
        }
    }
}

Function Clean-Up {
    $cDllFilepathList = @("C:\pcoip_perf_provider32.dll")
    foreach ($cDllFilepath in $cDllFilepathList) {
        if (Test-Path -Path $cDllFilepath) {
            Log-Info ("Deleting the $cDllFilepath file from C Drive");
            try {
                Remove-Item -Path $cDllFilepath -Force
            }
            catch {
                Log-Error ("Exception in deleting  dll file")
            }
        }
    }
    $oldUninstallerDirectory = "C:\Program Files (x86)\Teradici\PCoIP Agent\bin1x"
    if (Test-Path -Path $oldUninstallerDirectory) {
        Log-Info ("Removing bin1x")
        try {
            Remove-Tree $oldUninstallerDirectory
        }
        catch {
            Log-Error ("Exception in removing bin1x directory: ")
        }
    }
    $installRootOld = "C:\Program Files (x86)\Teradici.old"
    if (Test-Path -Path $installRootOld) {
        Log-Info ("Deleting Teradici.old directory")
        try {
            Remove-Item -Path $installRootOld -Force -Recurse
        }
        catch {
            Log-Error ("Exception in deleting Teradici.old installation directory")
        }
    }

}

Function Clean-Up-Teradici {
    $serviceNameList = Get-WmiObject -Class Win32_Service | Where-Object { $_.Name -like "PCoIP*" }
    foreach ($service in $serviceNameList) {
        if ($service) {
            Log-Info ("Deleting $service service")
            try {
                if ($service.State -eq "Running") {
                    $service.stopservice()
                }
                else {
                    Log-Info "$($service.name) is in $($service.State)"
                }
            }
            catch {
                Log-Error ("Exception in stopping $service service")
            }
            try {
                $service.delete()
            }
            catch {
                Log-Error ("Exception in deleting $service service")
            }
        }
        else {
            Log-Info "No PCoIP services found"
        }
    }

    $installRoot = "C:\Program Files (x86)\Teradici\"

    if (Test-Path -Path $installRoot) {
        Log-Info ("Renaming Teradici directory to Teradici.old")
        try {
            Rename-Item -Path $installRoot -NewName "Teradici.old" -Force
        }
        catch {
            Log-Error ("Exception in renaming Teradici.old installation directory")
        }
    }
    # Delete USB Hub drivers loaded by the Kernel using devcon util
    $application = "tera_dev_con"
    $manifestFileName = "${application}.zip.manifest.xml"
    $localManifestFile = ('{0}\{1}' -f $downloadDirectory, $manifestFileName)
    $remoteFileRootPath = "updates/apps/Utils/$application"
    Get-ManifestFile
    $packageFileName = "${application}.zip"
    $localPackageFile = ('{0}\{1}' -f $downloadDirectory, $packageFileName)
    $dsaSignature = Get-SignatureFromManifestFile $localManifestFile
    If (Test-Path -Path $localPackageFile) {
        Remove-Item -Path $localPackageFile -Force -ErrorAction SilentlyContinue
    }
    Download-FileFromUpdateServer -ServerFileRootPath $remoteFileRootPath -FileName $packageFileName -LocalFilePath $localPackageFile
    if (-not (Test-Path $localPackageFile)) {
        Log-Error ("Downloaded '{0}' package file '{1}' can't be found unexpectedly." -f $application, $localPackageFile)
        return
    }
    Log-Info ("Verify signature of downloaded '{0}' package file '{1}' against WorkSpaces certificate." -f $application, $localPackageFile)
    $workspacesCertThumbprint = Get-WorkSpacesCertThumbprint
    $verified = Verify-PackageFileSignature -PackageFilePath $localPackageFile -Signature $dsaSignature -CertThumbprint $workspacesCertThumbprint
    if (-not $verified) {
        Log-Error ("Downloaded package file '{0}' failed signature validation. Stop updating '{1}'." -f $localPackageFile, $application)
        Remove-Item $localPackageFile -Force -ErrorAction SilentlyContinue
        return
    }
    Log-Info ("Extract package '{0}' of '{1}' into '{2}'." -f $localPackageFile, $application, $downloadDirectory)
    Extract-Zip $localPackageFile $downloadDirectory

    $HardWareIDList = @("usbtstub", "vuhub")
    Log-Info ("Deleting device driver with $application")
    foreach ($HardWareID in $HardWareIDList) {
        $oemID = (Get-WmiObject -Class Win32_PnpSignedDriver | Where-Object { $_.hardwareID -eq $HardWareID } | Select-Object InfName -ExpandProperty InfName | Select-Object -Unique)
        if ($oemID) {
            Invoke-Expression "& '${downloadDirectory}\${application}.exe' remove $HardWareID"
            $Command = "& '${downloadDirectory}\${application}.exe' dp_delete $oemID -f"
            Log-Info ("Executing $Command")
            Invoke-Expression $Command
        }
        else { Log-Info "Could not find oem.inf file for $HardWareID . Skipping driver removal" }
    }

    # Delete PCoIP related regkey
    $reg_clean_list = @(
        "HKLM:\SOFTWARE\Teradici",
        "HKLM:\SOFTWARE\Wow6432Node\Teradici",
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\pcoip_agent.exe",
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\PCoIP Agent",
        "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\PCoIP Standard Agent",
        "HKLM:\SOFTWARE\Teradici\PCoIPAgent\DetectIdleShutdownAgentInstall",
        "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\PCoIP Graphics Agent")

    foreach ($regkey in $reg_clean_list) {
        if (Test-Path -Path $regkey) {
            Log-Info ("Deleting {0} entry" -f $regkey);
            try {
                Remove-Item -Path $regkey -Force -Recurse
            }
            catch {
                Log-Error ("Exception in deleting registry {0}" -f $regkey)
            }
        }
    }
}

Function Remove-OldTeradiciLogFile {
    param (
        [Parameter(Mandatory)]$path,
        [Parameter(Mandatory)]
        [Int]$olderThanInDays
    )
    Log-Info "Attempting to remove all logs older than $olderThanInDays day(s) in $path"
    try {
        Get-ChildItem $path -Recurse -Force -Filter *.log |  Where-Object { -not$_.PSIsContainer -and $_.CreationTime -le ((Get-Date).AddDays(-$OlderThanInDays)) } | Remove-Item -Force
        Log-Info "Successfully removed logs from $path"
    }
    catch {
        Log-Exception $_
        Log-Debug "Unable to remove logs in $path"
    }
}

Function Update-PCoIP_Update_state {
    Param ($PCoIPUpdate)
    Set-RegistryKeyValue -KeyPath $BackupPath -KeyName $PCoIPAgentStatusKeyName -PropertyType 'String' -Value $PCoIPUpdate
    Log-Info "Setting PCoIP Agent Status in .applications file status: $PCoIPUpdate"
    Update-RegisteredSkyLightApplicationVersion -ApplicationName "$PCoIPAgentStatusKeyName" -Version $PCoIPUpdate | Out-Null
}

Function Add-PCoIP_Legacy_regkey {
    # Ensure we have the logpuller directories configured appropriately - agents > RT4(?) use reg key in Wow6432Node path, so we need to create the legacy path because the Skylight agent looks there and not in the new path
    $PCoIPlogpath = "C:\ProgramData\Teradici\PCoIPAgent\logs"
    $LegacyTeradiciRegPath = "HKLM:\SOFTWARE\Teradici"
    $LegacyTeradiciRegKey = "PCoIPAgent"
    if (-not(Get-ItemProperty $("$LegacyTeradiciRegPath\$LegacyTeradiciRegKey") -Name LogPath -ErrorAction SilentlyContinue)) {
        Log-Info "Missing legacy registry key for PCoIP log directory, creating it..."
        try {
            # Check for parent path HKLM:\SOFTWARE\Teradici\PCoIPAgent, if not create one. Force creating will removes the existing dependent subkeys.
            if (-not(Test-Path -Path "$LegacyTeradiciRegPath\$LegacyTeradiciRegKey")) {
                Log-Info "Missing legacy registry key for PCoIP log directory at '$LegacyTeradiciRegPath\$LegacyTeradiciRegKey'. Proceeding to create."
                New-Item -Path $LegacyTeradiciRegPath -Name $LegacyTeradiciRegKey -Force
            }
            New-ItemProperty -Path $("$LegacyTeradiciRegPath\$LegacyTeradiciRegKey") -Name LogPath -PropertyType String -Value $PCoIPlogpath -Force
            Log-Info "Created legacy registry key for PCoIP log directory at '$LegacyTeradiciRegPath\$LegacyTeradiciRegKey' with value '$PCoIPlogpath'"

            # Now we need to see if we need to restart the Skylight agent to pick up the change
            $configurationStatus = Get-SkyLightWorkspaceConfigServiceStatus
            Log-Info "WorkSpace Configuration state: '$configurationStatus'"
            if (($configurationStatus -eq 2) -and ($(Get-Service -Name SkyLightWorkspaceConfigService).Status -eq "Running")) {
                Log-Info "WorkSpace Configuration is complete and the Skylight agent is running - restarting the Skylight agent to pick up the new logpath value"

                try {
                    Restart-Service -Name SkyLightWorkspaceConfigService -Force
                    Log-Info "Successfully restarted the Skylight agent service"
                }
                catch {
                    Log-Error "Error encountered restarting Skylight service"
                    throw $_
                }
            }
            else {
                Log-Info "WorkSpace configuration is not yet complete or the Skylight agent is not yet running - no need to restart the Skylight agent"
            }
        }
        catch {
            Log-Error "Failure encountered attempting to create legacy PCoIP log key"
            Log-Exception $_
        }
    }
    else {
        Log-Info "Legacy PCoIP log path reg key exists, skip creating the key/value"
    }
}

Function Set-ProcessDumpRegKey {
    $ParentPath = "HKLM:\SOFTWARE\Microsoft\Windows\Windows Error Reporting\LocalDumps"
    $ChildNames = @("pcoip_agent.exe", "pcoip_arbiter_win32.exe", "pcoip_vchan_loader.exe")
    $DumpKey = "DumpCount"
    ForEach ($Child in $ChildNames) {
        Set-RegistryKeyValue -KeyPath "$ParentPath\$Child" -KeyName $DumpKey -PropertyType "DWord" -Value 0 -Exit
    }
}

Function Backup-UsbRedirectionSetting {
    #Backup only if USB redirection is enabled, pcoip.enable_usb == 1
    $CurrentSettingRedirection = Get-RegistryKeyValue -KeyPath $UsbRedirectionKeyPath -KeyName $UsbRedirectionKeyName
    $CurrentSettingAuth = Get-RegistryKeyValue -KeyPath $UsbRedirectionKeyPath -KeyName $UsbAuthTableKeyName
    $CurrentSettingRedirection_adm = Get-RegistryKeyValue -KeyPath $UsbRedirectionKeyPath_NonAdm -KeyName $UsbRedirectionKeyName
    $CurrentSettingAuth_adm = Get-RegistryKeyValue -KeyPath $UsbRedirectionKeyPath_NonAdm -KeyName $UsbAuthTableKeyName
    ForEach ($Key in @("$UsbRedirectionKeyName", "$UsbAuthTableKeyName", "$UsbAdmKeyName")) {
        Remove-RegistryKeyValue -KeyPath "$BackupPath" -KeyName $Key
    }
    #Overridable settings in pcoip_adm takes preference over pcoip_admin_defaults. In case of a conflict in config, pcoip_admin takes preference.
    If ($CurrentSettingRedirection -eq 1 -and $Null -ne $CurrentSettingAuth) {
        Set-RegistryKeyValue -KeyPath $BackupPath -KeyName $UsbAdmKeyName -PropertyType 'String' -Value 'True' -Exit
        Set-RegistryKeyValue -KeyPath $BackupPath -KeyName $UsbRedirectionKeyName -PropertyType 'DWord' -Value $CurrentSettingRedirection -Exit
        Set-RegistryKeyValue -KeyPath $BackupPath -KeyName $UsbAuthTableKeyName -PropertyType 'String' -Value $CurrentSettingAuth -Exit
    }
    ElseIf ($CurrentSettingRedirection_adm -eq 1 -and $Null -ne $CurrentSettingAuth_adm) {
        Set-RegistryKeyValue -KeyPath $BackupPath -KeyName $UsbAdmKeyName -PropertyType 'String' -Value 'False' -Exit
        Set-RegistryKeyValue -KeyPath $BackupPath -KeyName $UsbRedirectionKeyName -PropertyType 'DWord' -Value $CurrentSettingRedirection_adm -Exit
        Set-RegistryKeyValue -KeyPath $BackupPath -KeyName $UsbAuthTableKeyName -PropertyType 'String' -Value $CurrentSettingAuth_adm -Exit
    }
    Else {
        Log-Info -Message "No valid combination for USB redirection found at location '$UsbRedirectionKeyPath' or '$UsbRedirectionKeyPath_NonAdm'"
    }
}

Function Set-UsbRedirectionSetting {
    $BackupSettingRedirection = Get-RegistryKeyValue -KeyPath $BackupPath -KeyName $UsbRedirectionKeyName
    $BackupSettingAuth = Get-RegistryKeyValue -KeyPath $BackupPath -KeyName $UsbAuthTableKeyName
    $BackupSettingAdmKey = Get-RegistryKeyValue -KeyPath $BackupPath -KeyName $UsbAdmKeyName
    If ($BackupSettingRedirection -eq 1 -and $Null -ne $BackupSettingAuth) {
        If ('True' -eq $BackupSettingAdmKey) {
            Set-RegistryKeyValue -KeyPath $UsbRedirectionKeyPath -KeyName $UsbRedirectionKeyName -PropertyType 'DWord' -Value $BackupSettingRedirection -Exit
            Set-RegistryKeyValue -KeyPath $UsbRedirectionKeyPath -KeyName $UsbAuthTableKeyName -PropertyType 'String' -Value $BackupSettingAuth -Exit
        }
        Else {
            Set-RegistryKeyValue -KeyPath $UsbRedirectionKeyPath_NonAdm -KeyName $UsbRedirectionKeyName -PropertyType 'DWord' -Value $BackupSettingRedirection -Exit
            Set-RegistryKeyValue -KeyPath $UsbRedirectionKeyPath_NonAdm -KeyName $UsbAuthTableKeyName -PropertyType 'String' -Value $BackupSettingAuth -Exit
        }
    }
    Else {
        Set-RegistryKeyValue -KeyPath $UsbRedirectionKeyPath -KeyName $UsbRedirectionKeyName -PropertyType 'DWord' -Value 0 -Exit
        Set-RegistryKeyValue -KeyPath $UsbRedirectionKeyPath_NonAdm -KeyName $UsbRedirectionKeyName -PropertyType 'DWord' -Value 0 -Exit
    }
}

Function Disable-PCoIP {
    #Disable Services
    $DriversToDisable = Get-WmiObject -ClassName Win32_SystemDriver | Where-Object { $_.DisplayName -in ("Service for Teradici Virtual Audio Driver", "TeraKDOD", "Teradici VHID", "Teradici USB Stub", "Teradici Virtual USB Hub") }
    Get-PnpDevice | Where-Object { $_.FriendlyName -match "Teradici" } | Disable-PnpDevice -Confirm:$false -ErrorAction SilentlyContinue
    $Win32_Service = Get-Service -Name PCoIPAgent, PCoIPPrintingSvc
    ForEach ($Driver in $DriversToDisable) {
        Log-Info "The driver '$($Driver.Name)' has a present startmode set to '$($driver.StartMode)' and present state is '$($Driver.Status)'. Stopping driver and setting startmode to disabled."
        #Disabling the driver serivce would prevent it from getting into running state post next reboot.
        $Driver.ChangeStartMode("Disabled")
    }
    # Disabling PCoIP recovery by Skylight since the function will follow reboot, we are not not restarting skylight
    $SkylightserviceState = Get-Service SkyLightWorkspaceConfigService
    $SkylightserviceState | Stop-Service -Force
    Log-Info "Stopped Skylight service $(Get-Service SkyLightWorkspaceConfigService)"
    $PCoIPRecoveryBySkylightKeyPath = "HKLM:\SOFTWARE\Amazon\Skylight\ConfigurationData"
    Log-Info "Adding registry key PCoIPRecoveryDisabled = true to disable start of PCoIP by Skylight"
    Set-RegistryKeyValue -KeyPath $PCoIPRecoveryBySkylightKeyPath -KeyName "PCoIPRecoveryDisabled" -PropertyType String -Value 'True'
    If ($Null -ne $Win32_Service) {
        Log-Info "The service '$($Win32_Service.Name)' has a present status as '$($Win32_Service.Status)'. Stopping service and setting Startup Type to Disabled."
        $Win32_service | Stop-Service
        $Win32_Service | Set-Service -Status Stopped -StartupType Disabled
        Log-Info "Current status of service '$($Win32_Service.Name)' is '$($Win32_Service.Status)'."
    }
    $controller = Get-WmiObject -Class Win32_VideoController
    Disable-PnpDevice -InstanceId $($controller.PNPDeviceID) -Confirm:$false
    Update-RegisteredSkyLightApplicationVersion -ApplicationName "PcoipStatus" -Version "Disabled" | Out-Null

    #Disabling scheduled task for migrating to WSPV2
    $ExistingPCOIPMajorVersion = ($Script:ExistingPCoIPVersion).Major
    Log-info -Message "Existing pcoip version is $Script:ExistingPCoIPVersion"
    if ($ExistingPCOIPMajorVersion -ge 20) {
        Log-Info "Disabling TeraRestoreWDDMDriver scheduled task. "
        Start-Process "c:\windows\system32\schtasks.exe" -ArgumentList "/End /TN $Script:TaskName"
        Start-Process "c:\windows\system32\schtasks.exe" -ArgumentList "/change /TN $Script:TaskName /Disable"
    }
    else {
        Log-Info "Skip disabling TeraRestoreWDDMDriver scheduled task since the PCOIP major build is $ExistingPCOIPMajorVersion"
    }
}

Function Enable-PCoIP {
    #Enable Services
    $DriversToEnable = Get-WmiObject -ClassName Win32_SystemDriver | Where-Object { $_.DisplayName -in ("Service for Teradici Virtual Audio Driver", "TeraKDOD", "Teradici VHID", "Teradici USB Stub", "Teradici Virtual USB Hub") }
    Get-PnpDevice | Where-Object { $_.FriendlyName -match "Teradici" } | Enable-PnpDevice -Confirm:$false -ErrorAction SilentlyContinue
    Log-Info "Querying Details of all the services in the WorkSpace"
    $Win32_Service = Get-Service -Name PCoIPAgent, PCoIPPrintingSvc
    ForEach ($Driver in $DriversToEnable) {
        Log-Info "The driver '$($Driver.Name)' has a present startmode set to '$($driver.StartMode)' and present state is '$($Driver.Status)'. Starting driver and setting startmode to manual."
        $Driver.ChangeStartMode("manual")
    }
    # Enabling start of PCoIP by Skylight
    $PCoIPRecoveryBySkylightKeyPath = "HKLM:\SOFTWARE\Amazon\Skylight\ConfigurationData"
    Log-Info "Removing registry key PCoIPRecoveryDisabled = true to enable start of PCoIP by Skylight"
    Remove-RegistryKeyValue -KeyPath $PCoIPRecoveryBySkylightKeyPath -KeyName "PCoIPRecoveryDisabled"
    If ($Null -ne $Win32_Service) {
        Log-Info "The service '$($Win32_Service.Name)' has a present status as '$($Win32_Service.Status)'. Starting service and setting Startup Type to Automatic."
        $Win32_Service | Set-Service -Status Running -StartupType Automatic
        $Win32_service | Start-Service
        Log-Info "Current status of service '$($Win32_Service.Name)' is '$($Win32_Service.Status)'."
    }
    $controller = Get-WmiObject -Class Win32_VideoController
    Enable-PnpDevice -InstanceId $($controller.PNPDeviceID) -Confirm:$false
    Update-RegisteredSkyLightApplicationVersion -ApplicationName "PcoipStatus" -Version "Enabled" | Out-Null

    #Enabling scheduled task for migrating to WSPV2
    $ExistingWSPMajorVersion = ($Script:ExistingWSPVersion).Major
    $ExistingPCOIPMajorVersion = ($Script:ExistingPCoIPVersion).Major
    Log-info -Message "Existing WSP version is $Script:ExistingWSPVersion and existing PCoIP version is $Script:ExistingPCoIPVersion"
    if ($ExistingWSPMajorVersion -ge 2) {
        If ($ExistingPCOIPMajorVersion -ge 20) {
            Log-info -Message "Enabling and starting TeraRestoreWDDMDriver scheduled task."
            Start-Process "c:\windows\system32\schtasks.exe" -ArgumentList "/change /TN $Script:TaskName /Enable"
            Start-Process "c:\windows\system32\schtasks.exe" -ArgumentList "/Run /TN $Script:TaskName"
        }
        Else {
            Log-info -Message "Using DevCon.exe util to associate display adapter with tera_kmdod.sys"
            $infName = ((. c:\windows\system32\pnputil.exe -e | Select-String -Context 2 'Class :\s+ Display' | Where-Object { $_ -match 'Teradici' }).Context.PreContext)[0].Split(':')[1].Trim()
            Log-Info "Inf name for TeraKDOD is '$infName'"
            & 'C:\Program Files\NICE\DCV\Server\drivers\common\devcon.exe' updateni "C:\Windows\INF\$infName" 'pci\cc_0300'
            Log-info -Message "Restarting the WorkSpace after updating the display driver."
            Restart-WksNow
        }
    }
}

Function Backup-PCoIPRegKey {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $false)]
        [String]$ExportPathFile = $Script:downloadDirectory
    )
    $keys = 'HKLM:\SOFTWARE\WOW6432Node\Policies\Teradici', 'HKLM:\SOFTWARE\Policies\Teradici', 'HKLM:\SOFTWARE\Teradici\PCoIP'
    Log-Info ("Backing up $keys registry")
    Try {
        $OutputFile = "$($Script:downloadDirectory)\PCoIPRegKeyBackup.reg"
        Get-ChildItem -Path $ExportPathFile -Filter '*.reg' | Remove-Item
        $keys |
        ForEach-Object {
            $RegFileName++
            $RegInfo = Get-ChildItem -Path $_
            If ($RegInfo) {
                Invoke-Expression "cmd.exe /c reg export $RegInfo '$ExportPathFile\$RegFileName.reg' /y"
                If(Test-Path -Path $ExportPathFile -Filter '*.reg'){
                Log-info -Message "Registry Backup was Sucessful"
                }
                Else{
                    Log-info -Message "Registry Backup was Unsucessful"
                }
            }
        }
        Get-Content -Path "$ExportPathFile\*.reg" | Set-Content $OutputFile
        }
    Catch {
        Log-Exception $_
    }
}

Function Restore-PCoIPRegKey {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RegKeyPath
    )
    Try {
        $argList = '/s {0}' -f $RegKeyPath
        Start-Process -FilePath "$env:windir\regedit.exe" -ArgumentList $argList
        Log-Info "Restore RegKeys for PCoIPRegKeyBackup was successful"
    }
    Catch {
        Log-Exception $_
    }
}

Function Disable-USBWebcam {
    Foreach ($DisableUsbWebcam in $Script:DisablePCoIPWebcamRegKeys) {
        $MetricName = $DisableUsbWebcam.split('\')[-1] + '_enable_usb_video'
        Log-Info ("Disabling USBWebcam $DisableUsbWebcam registry")
        Try {
            Set-RegistryKeyValue -KeyPath $DisableUsbWebcam -KeyName "pcoip.enable_usb_video" -PropertyType 'Dword' -Value '0' -Exit
            Update-RegisteredSkyLightApplicationVersion -ApplicationName $MetricName -Version "0" | Out-Null
            Log-Info -Message "Disabled USB Webcam at RegKeyPath:'$DisableUsbWebcam'"
        }
        Catch {
            Log-Exception $_
        }
    }
}

Function Test-ShouldInstall {
    <#
    .SYNOPSIS
    This function determines the elibility to install WSP.
    #>
    Param(
        [Parameter(Mandatory = $false)]$WspApplicable
    )
    Try {
        #Checks if InstallType entry value is PcoipWspAllBundles in wsp manifest file
        If ($Script:WspApplicable.InstallType -eq "PcoipWspAllBundles") { return $true }
        #Checks if AppEnabledKeyStat entry in registries
        If (-not($Script:WspApplicable.AppEnabledKeyStat)) {
            #Checks the InstallType entry value in wsp manifest file to see if anything is matched in allowedstate list
            If ($Script:WspApplicable.InstallType -in $Script:AllowedState) { return $true }
        }
        Else {
            return $false
        }
        return $true
    }
    Catch {
        # Do nothing if there is an exception
        Log-Exception $_
        return $false
    }
}

Function Search-LogonBanner {
    # This function searches through certain registry keys to identify if there is any value for the logon banner and sends telemetry information.
    [string]$Result = "False"
    $RegPaths = @(
        "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon",
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System"
    )
    $RegKeys = @("LegalNoticeCaption", "LegalNoticeText")
    try {
        foreach ($RegPath in $RegPaths) {
            foreach ($RegKey in $RegKeys) {
                $RegValue = Get-RegistryKeyValue -KeyPath $RegPath -KeyName $RegKey
                if (("" -eq ("{0}" -f $RegValue).Trim())) {
                    $Result = "False"
                    Log-Info "LogonBanner not Detetced"
                }
                else {
                    $Result = "True"
                    Log-Info "LogonBanner Detetced"
                    break
                }
            }
        }
    }
    catch {
        Log-Exception $_
        $Result = "Error"
    }
    finally {
        $Script:LogOnBannerExists = $Result
        Update-RegisteredSkyLightApplicationVersion -ApplicationName "LogonBanner" -Version $Result | Out-Null
    }
}

Function Execute-Function {

    Set-ScriptVariable
    Init-Logging
    try {
        Init-WorkSpacesConfigRegistryKey
        Set-WorkSpacesConfigScriptVersionInRegistry -ScriptName $scriptName -Version $scriptVersion
    }
    catch {
        Log-Exception $_
    }
    try {
        Log-Info ("Script version date: $scriptVersion")
        Log-Info ("Check whether the Windows version is supported by '{0}'" -f $application)
        $windowsVersionSupported = Is-SupportedWindowsVersion
        if (-not $windowsVersionSupported[0]) {
            Log-Info ("The Windows version is not supported by '{0}'. Skip checking update and configuring '{0}'." -f $application)
            return
        }
        if (-not(Is-DomainJoinFinished)) {
            Log-Info -Message "Workspace is not domain joined. Exiting $scriptName execution"
            Exit;
        }
    }
    catch {
        Log-Exception $_
    }
    Search-LogonBanner
    try {
        # Block for handling Seamless Migration request
        # Get WorkSpace Config File
        $Script:WorkspaceConfig = Get-WorkspaceConfigFileFromUpdateServer -WorkspaceConfigFileName $WorkspaceConfigFileName -LocalTempWorkspaceConfigFilePath $LocalTempWorkspaceConfigFilePath -WorkspaceConfigRemoteFileRootPath $WorkspaceConfigRemoteFileRootPath -UpdateServerIpAddress (Get-UpdateServerIp)
        $Script:PCoIPServiceExist = Get-Service -Name PCoIPAgent -ErrorAction SilentlyContinue
        if ($null -eq $Script:PCoIPServiceExist) {
            # PCoIP Service doesnot Exist
            Log-Info -Message "PCoIP Service does not exist"
        }
        Else {
            # PCoIP Service Exist. Get the Service status
            Log-Info -Message "PCoIP Service exists. Querying Service status"
            $Script:PCoIPAgentStatus = (Get-Service PCoIPAgent).Status
            $Script:PCoIPAgentStartType = (Get-Service PCoIPAgent).StartType
            Log-Info -Message ("PCoIP agent service status: $PCoIPAgentStatus and StartType is $PCoIPAgentStartType ")
        }
        If ($WorkspaceConfig) {
            Log-Info -Message "Protocol Migration Of WorkSpace is requested"
            # Get the content of WorkSpace Config File for Migration
            $Script:WorkspaceConfigJsonContents = Get-WorkspaceConfigJsonInformation -LocalTempWorkspaceConfigFilePath $LocalTempWorkspaceConfigFilePath -Application $ConfigFileName
            If ($Script:WorkspaceConfigJsonContents.ProtocolType -eq "wsp" ) {
                Set-WorkSpaceProtocolMigration -PreferredProtocol "WSP"
                Log-Info -Message "The WorkSpace has requested to migrate to wsp, thus PCoIP needs to be disabled"
                Log-Info -Message "Proceeding with disabling PCoIP"
                If ($null -eq $Script:PCoIPServiceExist -or ($Script:PCoIPAgentStatus -eq "Stopped" -and $PCoIPAgentStartType -eq "Disabled")) {
                    Log-Info -Message "PCoIP service does not exist or is already disabled. Exiting other checks for PCoIP"
                    Exit
                }
                Else {
                    Disable-PCoIP
                    Log-Info -Message "Reboot is required post stopping of PCoIP Agents and Drivers."
                    Set-ProtocolMigrationStatus -ProtocolMigrationStatus "PCOIP_Disabled"
                    Restart-WKSNow -AppName "PCoIPDisabledForMigration"
                }
            }
            Elseif ($Script:WorkspaceConfigJsonContents.ProtocolType -eq "pcoip") {
                Set-WorkSpaceProtocolMigration -PreferredProtocol "PCOIP"
                Log-Info -Message "The WorkSpace has requested to migrate to PCoIP, thus PCoIP needs to be enabled"
                If ($null -eq $Script:PCoIPServiceExist) {
                    Log-Info -Message "PCoIP Service does not exist and will be installed now"
                }
                ElseIf ($Script:PCoIPAgentStatus -eq "Running") {
                    Log-Info -Message "PCoIP is already Running."
                    Set-ProtocolMigrationStatus -ProtocolMigrationStatus "PCOIP_MigrationSuccessful"
                }
                Else {
                    Log-Info -Message "Proceeding with enabling PCoIP"
                    Set-ProtocolMigrationStatus -ProtocolMigrationStatus "PCOIP_InProgress"
                    Enable-PCoIP
                    If ($Script:PCoIPAgentStatus -eq "Running") {
                        Log-Info -Message "Migration to pcoip is successful. Reboot is required post starting of PCoIP Agents and Drivers."
                        Set-ProtocolMigrationStatus -ProtocolMigrationStatus "PCOIP_MigrationSuccessful"
                        Restart-WKSNow -AppName "PCoIPEnabledForMigration"
                    }
                    Else {
                        Log-Info -Message "Migration to pcoip failed."
                        Set-ProtocolMigrationStatus -ProtocolMigrationStatus "PCOIP_MigrationFailed"
                    }
                }
            }
            Else {
                Set-WorkSpaceProtocolMigration -PreferredProtocol "UNKNOWN"
                Log-Info -Message "Workspace_config.json was detected but no valid protocol migration is present in config file"
            }
        }
        Else {
            Log-Info -Message "Protocol migration of the workspace is not requested"
        }
        #PCoIP should be installed on machines if InstallType = NULL or "StxhdDisabledWspInstalled","StxhdRemovedWspInstalled" & Wsp is not applicable
        #PCoIP should be installed on machines if InstallType = PcoipWspAllBundles
        $ShouldInstall = Test-ShouldInstall -WspApplicable $Script:WspApplicable
        Log-Info -Message "'ShouldInstall' value is '$ShouldInstall'. WSP AppApplicable value is '$($Script:WspApplicable.AppApplicable)', Wsp InstallType value is '$($Script:WspApplicable.InstallType)', Wsp InstallType Length is '$($Script:WspApplicable.InstallType.Length)' and AppEnabledKeyStat value is '$($Script:WspApplicable.AppEnabledKeyStat)'"
        $UserLoggedOn = (Get-WsOsInfo).UserActive
        If ($UserLoggedOn) {
            Log-Info -Message "Users found logged in to the workspace"
            Log-Info -Message "Exiting $ScriptName"
            Exit
        }
        Else {
            Log-Info -Message "No users found logged in to the workspace."
        }
        $Script:ExistingVersion = Get-InstalledPCoIPVersion
        If (-not $ShouldInstall) {
            Log-Info "Checking if PCoIP is installed and removing if necessary"
            If ($Null -ne $ExistingVersion -and 'NULL' -ne $ExistingVersion) {
                Log-Info "Existing PCoIP installation found, version '$ExistingVersion'"
                If ($ExistingVersion -lt '20.10.4') {
                    $oldUninstallerDirectory = "C:\Program Files (x86)\Teradici\PCoIP Agent"
                    $oldUninstallerFile = "C:\Program Files (x86)\Teradici\PCoIP Agent\uninst.exe"
                }
                Else {
                    $oldUninstallerDirectory = "C:\Program Files\Teradici\PCoIP Agent\PCoIP Agent"
                    $oldUninstallerFile = "C:\Program Files\Teradici\PCoIP Agent\uninst.exe"
                }
                #Attempt Uninstall
                $UninstallPCoIPResult = Uninstall-PCoIPAgent
                If ($UninstallPCoIPResult[0] -and $UninstallPCoIPResult[1]) {
                    Log-Info "Reboot is required post un-installation of PCoIP Agent on the Wsp Only box."
                    Restart-WKSNow -AppName "PCoIPRebootStatus"
                }
            }
            Else {
                Log-Info -Message "PCoIP is not installed on the box."
            }
        }
        Else {
            Log-Info ("Cleaning up before updating")
            Clean-Up
            Log-Info ("Check update for '{0}'." -f $application)
            Update-PCoIPAgent
            Log-Info ("Finish checking update for '{0}'." -f $application)
            Log-Info ("Configure '{0}'." -f $application)
            Disable-USBWebcam
            Configure-PCoIPAgent
            Log-Info ("Finish configuring '{0}'." -f $application)
        }
    }
    catch {
        Update-RegisteredSkyLightApplicationVersion -ApplicationName "PCoIP_Agent_Update_Failure" -Version "Unknown" | Out-Null
        Log-Exception $_
    }
    finally {
        Log-Info ("Finish checking update and configuring for '{0}'." -f $application)
    }
}

Function Set-Main {
    Execute-Function
}

Set-Main

# SIG # Begin signature block
# MIIuBgYJKoZIhvcNAQcCoIIt9zCCLfMCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCA+qoYV5qj45XNS
# alNVV0X7wqCttBwyCidguOoMnFe/waCCE3MwggXAMIIEqKADAgECAhAP0bvKeWvX
# +N1MguEKmpYxMA0GCSqGSIb3DQEBCwUAMGwxCzAJBgNVBAYTAlVTMRUwEwYDVQQK
# EwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xKzApBgNV
# BAMTIkRpZ2lDZXJ0IEhpZ2ggQXNzdXJhbmNlIEVWIFJvb3QgQ0EwHhcNMjIwMTEz
# MDAwMDAwWhcNMzExMTA5MjM1OTU5WjBiMQswCQYDVQQGEwJVUzEVMBMGA1UEChMM
# RGlnaUNlcnQgSW5jMRkwFwYDVQQLExB3d3cuZGlnaWNlcnQuY29tMSEwHwYDVQQD
# ExhEaWdpQ2VydCBUcnVzdGVkIFJvb3QgRzQwggIiMA0GCSqGSIb3DQEBAQUAA4IC
# DwAwggIKAoICAQC/5pBzaN675F1KPDAiMGkz7MKnJS7JIT3yithZwuEppz1Yq3aa
# za57G4QNxDAf8xukOBbrVsaXbR2rsnnyyhHS5F/WBTxSD1Ifxp4VpX6+n6lXFllV
# cq9ok3DCsrp1mWpzMpTREEQQLt+C8weE5nQ7bXHiLQwb7iDVySAdYyktzuxeTsiT
# +CFhmzTrBcZe7FsavOvJz82sNEBfsXpm7nfISKhmV1efVFiODCu3T6cw2Vbuyntd
# 463JT17lNecxy9qTXtyOj4DatpGYQJB5w3jHtrHEtWoYOAMQjdjUN6QuBX2I9YI+
# EJFwq1WCQTLX2wRzKm6RAXwhTNS8rhsDdV14Ztk6MUSaM0C/CNdaSaTC5qmgZ92k
# J7yhTzm1EVgX9yRcRo9k98FpiHaYdj1ZXUJ2h4mXaXpI8OCiEhtmmnTK3kse5w5j
# rubU75KSOp493ADkRSWJtppEGSt+wJS00mFt6zPZxd9LBADMfRyVw4/3IbKyEbe7
# f/LVjHAsQWCqsWMYRJUadmJ+9oCw++hkpjPRiQfhvbfmQ6QYuKZ3AeEPlAwhHbJU
# KSWJbOUOUlFHdL4mrLZBdd56rF+NP8m800ERElvlEFDrMcXKchYiCd98THU/Y+wh
# X8QgUWtvsauGi0/C1kVfnSD8oR7FwI+isX4KJpn15GkvmB0t9dmpsh3lGwIDAQAB
# o4IBZjCCAWIwDwYDVR0TAQH/BAUwAwEB/zAdBgNVHQ4EFgQU7NfjgtJxXWRM3y5n
# P+e6mK4cD08wHwYDVR0jBBgwFoAUsT7DaQP4v0cB1JgmGggC72NkK8MwDgYDVR0P
# AQH/BAQDAgGGMBMGA1UdJQQMMAoGCCsGAQUFBwMDMH8GCCsGAQUFBwEBBHMwcTAk
# BggrBgEFBQcwAYYYaHR0cDovL29jc3AuZGlnaWNlcnQuY29tMEkGCCsGAQUFBzAC
# hj1odHRwOi8vY2FjZXJ0cy5kaWdpY2VydC5jb20vRGlnaUNlcnRIaWdoQXNzdXJh
# bmNlRVZSb290Q0EuY3J0MEsGA1UdHwREMEIwQKA+oDyGOmh0dHA6Ly9jcmwzLmRp
# Z2ljZXJ0LmNvbS9EaWdpQ2VydEhpZ2hBc3N1cmFuY2VFVlJvb3RDQS5jcmwwHAYD
# VR0gBBUwEzAHBgVngQwBAzAIBgZngQwBBAEwDQYJKoZIhvcNAQELBQADggEBAEHx
# qRH0DxNHecllao3A7pgEpMbjDPKisedfYk/ak1k2zfIe4R7sD+EbP5HU5A/C5pg0
# /xkPZigfT2IxpCrhKhO61z7H0ZL+q93fqpgzRh9Onr3g7QdG64AupP2uU7SkwaT1
# IY1rzAGt9Rnu15ClMlIr28xzDxj4+87eg3Gn77tRWwR2L62t0+od/P1Tk+WMieNg
# GbngLyOOLFxJy34riDkruQZhiPOuAnZ2dMFkkbiJUZflhX0901emWG4f7vtpYeJa
# 3Cgh6GO6Ps9W7Zrk9wXqyvPsEt84zdp7PiuTUy9cUQBY3pBIowrHC/Q7bVUx8ALM
# R3eWUaNetbxcyEMRoacwggawMIIEmKADAgECAhAIrUCyYNKcTJ9ezam9k67ZMA0G
# CSqGSIb3DQEBDAUAMGIxCzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2VydCBJ
# bmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xITAfBgNVBAMTGERpZ2lDZXJ0
# IFRydXN0ZWQgUm9vdCBHNDAeFw0yMTA0MjkwMDAwMDBaFw0zNjA0MjgyMzU5NTla
# MGkxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UE
# AxM4RGlnaUNlcnQgVHJ1c3RlZCBHNCBDb2RlIFNpZ25pbmcgUlNBNDA5NiBTSEEz
# ODQgMjAyMSBDQTEwggIiMA0GCSqGSIb3DQEBAQUAA4ICDwAwggIKAoICAQDVtC9C
# 0CiteLdd1TlZG7GIQvUzjOs9gZdwxbvEhSYwn6SOaNhc9es0JAfhS0/TeEP0F9ce
# 2vnS1WcaUk8OoVf8iJnBkcyBAz5NcCRks43iCH00fUyAVxJrQ5qZ8sU7H/Lvy0da
# E6ZMswEgJfMQ04uy+wjwiuCdCcBlp/qYgEk1hz1RGeiQIXhFLqGfLOEYwhrMxe6T
# SXBCMo/7xuoc82VokaJNTIIRSFJo3hC9FFdd6BgTZcV/sk+FLEikVoQ11vkunKoA
# FdE3/hoGlMJ8yOobMubKwvSnowMOdKWvObarYBLj6Na59zHh3K3kGKDYwSNHR7Oh
# D26jq22YBoMbt2pnLdK9RBqSEIGPsDsJ18ebMlrC/2pgVItJwZPt4bRc4G/rJvmM
# 1bL5OBDm6s6R9b7T+2+TYTRcvJNFKIM2KmYoX7BzzosmJQayg9Rc9hUZTO1i4F4z
# 8ujo7AqnsAMrkbI2eb73rQgedaZlzLvjSFDzd5Ea/ttQokbIYViY9XwCFjyDKK05
# huzUtw1T0PhH5nUwjewwk3YUpltLXXRhTT8SkXbev1jLchApQfDVxW0mdmgRQRNY
# mtwmKwH0iU1Z23jPgUo+QEdfyYFQc4UQIyFZYIpkVMHMIRroOBl8ZhzNeDhFMJlP
# /2NPTLuqDQhTQXxYPUez+rbsjDIJAsxsPAxWEQIDAQABo4IBWTCCAVUwEgYDVR0T
# AQH/BAgwBgEB/wIBADAdBgNVHQ4EFgQUaDfg67Y7+F8Rhvv+YXsIiGX0TkIwHwYD
# VR0jBBgwFoAU7NfjgtJxXWRM3y5nP+e6mK4cD08wDgYDVR0PAQH/BAQDAgGGMBMG
# A1UdJQQMMAoGCCsGAQUFBwMDMHcGCCsGAQUFBwEBBGswaTAkBggrBgEFBQcwAYYY
# aHR0cDovL29jc3AuZGlnaWNlcnQuY29tMEEGCCsGAQUFBzAChjVodHRwOi8vY2Fj
# ZXJ0cy5kaWdpY2VydC5jb20vRGlnaUNlcnRUcnVzdGVkUm9vdEc0LmNydDBDBgNV
# HR8EPDA6MDigNqA0hjJodHRwOi8vY3JsMy5kaWdpY2VydC5jb20vRGlnaUNlcnRU
# cnVzdGVkUm9vdEc0LmNybDAcBgNVHSAEFTATMAcGBWeBDAEDMAgGBmeBDAEEATAN
# BgkqhkiG9w0BAQwFAAOCAgEAOiNEPY0Idu6PvDqZ01bgAhql+Eg08yy25nRm95Ry
# sQDKr2wwJxMSnpBEn0v9nqN8JtU3vDpdSG2V1T9J9Ce7FoFFUP2cvbaF4HZ+N3HL
# IvdaqpDP9ZNq4+sg0dVQeYiaiorBtr2hSBh+3NiAGhEZGM1hmYFW9snjdufE5Btf
# Q/g+lP92OT2e1JnPSt0o618moZVYSNUa/tcnP/2Q0XaG3RywYFzzDaju4ImhvTnh
# OE7abrs2nfvlIVNaw8rpavGiPttDuDPITzgUkpn13c5UbdldAhQfQDN8A+KVssIh
# dXNSy0bYxDQcoqVLjc1vdjcshT8azibpGL6QB7BDf5WIIIJw8MzK7/0pNVwfiThV
# 9zeKiwmhywvpMRr/LhlcOXHhvpynCgbWJme3kuZOX956rEnPLqR0kq3bPKSchh/j
# wVYbKyP/j7XqiHtwa+aguv06P0WmxOgWkVKLQcBIhEuWTatEQOON8BUozu3xGFYH
# Ki8QxAwIZDwzj64ojDzLj4gLDb879M4ee47vtevLt/B3E+bnKD+sEq6lLyJsQfmC
# XBVmzGwOysWGw/YmMwwHS6DTBwJqakAwSEs0qFEgu60bhQjiWQ1tygVQK+pKHJ6l
# /aCnHwZ05/LWUpD9r4VIIflXO7ScA+2GRfS0YW6/aOImYIbqyK+p/pQd52MbOoZW
# eE4wggb3MIIE36ADAgECAhAEstNiv5tANt0f/Jc5YoS+MA0GCSqGSIb3DQEBCwUA
# MGkxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UE
# AxM4RGlnaUNlcnQgVHJ1c3RlZCBHNCBDb2RlIFNpZ25pbmcgUlNBNDA5NiBTSEEz
# ODQgMjAyMSBDQTEwHhcNMjMxMTIxMDAwMDAwWhcNMjQxMTIwMjM1OTU5WjB/MQsw
# CQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHU2VhdHRs
# ZTEZMBcGA1UEChMQQW1hem9uLmNvbSwgSW5jLjETMBEGA1UECxMKd29ya3NwYWNl
# czEZMBcGA1UEAxMQQW1hem9uLmNvbSwgSW5jLjCCAaIwDQYJKoZIhvcNAQEBBQAD
# ggGPADCCAYoCggGBALqzvPLNPRb9rs7PFxX8zKBtM6EaAM5gb8NBSmHHBwOzYr3a
# yy+u+8oe79l8YmIb7rtdCpSeYnAmPnLJiTDn8yS6z7N4hzEyQOXFyV/A2aOl8jhX
# dUvgbXGxEV8aIa5LJZdlCHqQmePBvlQAvNbpLW0yx4jgpZW7TBqy+17Hz8K8tccw
# GWO00Gz3dged92y4XuT7T4ckps6CQ/igBgB2N9284mZCtvLPSL34kd+3hS3D7DnR
# PxqyZ2MTqW4k5ph3Wp813AV9ju68DoraplKYM7m6ls3AnxpAmNKcZOaKOXsDqFEW
# PykXjzrR9bXBPzIKhyP1t8cLcMsMSmmJgetBjdLtl5+j7zNndRdk9HvcKC7zH6m1
# KgoPVVPWiojsUkQ3JE2ua8EkG0len9vYC8FFPI0rjag3A3singBygLvJTyuu8Wk5
# qLxuBfAW0brpb3ikSqSYSfHCH++k6QobfWqAZEmuQt4cklcvuyiNYiILY8bnOueR
# EGiiKhIoN9NXy/UbdQIDAQABo4ICAzCCAf8wHwYDVR0jBBgwFoAUaDfg67Y7+F8R
# hvv+YXsIiGX0TkIwHQYDVR0OBBYEFJIK7xBfSb0CZhV7RWordFZq3xQLMD4GA1Ud
# IAQ3MDUwMwYGZ4EMAQQBMCkwJwYIKwYBBQUHAgEWG2h0dHA6Ly93d3cuZGlnaWNl
# cnQuY29tL0NQUzAOBgNVHQ8BAf8EBAMCB4AwEwYDVR0lBAwwCgYIKwYBBQUHAwMw
# gbUGA1UdHwSBrTCBqjBToFGgT4ZNaHR0cDovL2NybDMuZGlnaWNlcnQuY29tL0Rp
# Z2lDZXJ0VHJ1c3RlZEc0Q29kZVNpZ25pbmdSU0E0MDk2U0hBMzg0MjAyMUNBMS5j
# cmwwU6BRoE+GTWh0dHA6Ly9jcmw0LmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0
# ZWRHNENvZGVTaWduaW5nUlNBNDA5NlNIQTM4NDIwMjFDQTEuY3JsMIGUBggrBgEF
# BQcBAQSBhzCBhDAkBggrBgEFBQcwAYYYaHR0cDovL29jc3AuZGlnaWNlcnQuY29t
# MFwGCCsGAQUFBzAChlBodHRwOi8vY2FjZXJ0cy5kaWdpY2VydC5jb20vRGlnaUNl
# cnRUcnVzdGVkRzRDb2RlU2lnbmluZ1JTQTQwOTZTSEEzODQyMDIxQ0ExLmNydDAJ
# BgNVHRMEAjAAMA0GCSqGSIb3DQEBCwUAA4ICAQCbx/nwj1esOGGpLH52oz6xPiGI
# cn62kSXw4qhXElN0YbskREFD61HQDwK1EhinXBvjL7Ir+xGodvR2IpPWIPb36ixy
# gCUSZDLZXK847qPbVjxb5ULu8yqJNQkFoz6KcrqLLnkyu/cg9VuC3wfbVBaR77bX
# X3KdOSUoDXiih9zSF0kg6YafxXeDyU7qjEZEyhWYQtahnc27hT4fz2pvghamhae3
# a9r+gIjSIPBryrkQiw2u6s+6dIKB1djKlYbFghkkjkg8KaEoCIlS2l8zGTzIc4er
# 1eU4p2Rd3HCPdXNxnrpJR2oFNeAHCHqlrEkIcRZV0Tz4Hf4jHy96oAYEldkvpeDl
# tVFtvpKiV6AGCPtuZMDLPMffZj42XZn7oOB+1WRtmfyO7+vQhjm7oATQNeIIm+XN
# whpPFqY08oQSUW4LQTSkeWLz579lubDfao1+Ta7kSWkMKw5fmjLzVW5C9L4qiZM7
# wA8fEaKbRMFfjFD3+F9YNorf1VZS8Yl3agB26VV7dVfR4qQtVJvAGs14Bxlfbhb4
# mIuwB1ZgjJvuqkERGYnMJ112o0zLLLoyVz2e39jOkJfdm6AnphtIXMXuLXy6tfeF
# 4nRCcypS5xVFaCIiAcVrtWC5CsYshkAux5Qm3LS9ZSRFMNjLnYDE/WsWyx8WFnwi
# iaU3ehJuEZFNjqZhhzGCGekwghnlAgEBMH0waTELMAkGA1UEBhMCVVMxFzAVBgNV
# BAoTDkRpZ2lDZXJ0LCBJbmMuMUEwPwYDVQQDEzhEaWdpQ2VydCBUcnVzdGVkIEc0
# IENvZGUgU2lnbmluZyBSU0E0MDk2IFNIQTM4NCAyMDIxIENBMQIQBLLTYr+bQDbd
# H/yXOWKEvjANBglghkgBZQMEAgEFAKB8MBAGCisGAQQBgjcCAQwxAjAAMBkGCSqG
# SIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisGAQQBgjcCAQsxDjAMBgorBgEEAYI3
# AgEVMC8GCSqGSIb3DQEJBDEiBCBetfpjGviZJenmP0vTj2jMsLJHkd7TV8ScW65p
# 3p758jANBgkqhkiG9w0BAQEFAASCAYCoPG90LC7mtbmW6LB1nk9uvMQfDb5nRa5Y
# wHbRY3FCQhm+RMRJShof9rwaadLQ70iNeBNiqXbPdm9I+hW1eZg56U4I3q508IPB
# jqD9dBbTEKYwrj3VodruAm7eFlQ2EV2C+n7xBr/LORjVtoo1ZGSXn6pv0J2C25wW
# gTb0hm5sMj9VVl2hR7ZDpQOpX8I07vg7w2DBZim9z96G7fZCL75Az/Du612midDL
# 51fbncjxj0GgIIsnqbIaF5wClqjKSieIIveMatm1QEEKZMoSDrYs7A53fcGZwCo9
# YugYYc7ssyUcZB3RB2MxwBzAHlKzstmBOCLLrEAUkbfqHwuGZWTwctEinjguV1Lu
# sRzLUL1v5L2iMfkqvukwEGJVprY4zD10HeDPhIWkJYUwqB833SlraQZtC/AjL6SD
# OPhqBBvLiur+ESetK9s1eZp8Q7ukHeIktjbiisdOoBfKxsFAhxzCKots8hBlSjs4
# 0C6mRMFQffNwbUZ7EgS+lLce4o/bQd+hghc/MIIXOwYKKwYBBAGCNwMDATGCFysw
# ghcnBgkqhkiG9w0BBwKgghcYMIIXFAIBAzEPMA0GCWCGSAFlAwQCAQUAMHcGCyqG
# SIb3DQEJEAEEoGgEZjBkAgEBBglghkgBhv1sBwEwMTANBglghkgBZQMEAgEFAAQg
# uu9drFI6HTUAMV6OSuRoj4k/hCi9Lj3HxLgdxCtwMpUCEDEZJSSulBc2WP5r8gLc
# Sv8YDzIwMjQwMjI5MDEwNzEyWqCCEwkwggbCMIIEqqADAgECAhAFRK/zlJ0IOaa/
# 2z9f5WEWMA0GCSqGSIb3DQEBCwUAMGMxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5E
# aWdpQ2VydCwgSW5jLjE7MDkGA1UEAxMyRGlnaUNlcnQgVHJ1c3RlZCBHNCBSU0E0
# MDk2IFNIQTI1NiBUaW1lU3RhbXBpbmcgQ0EwHhcNMjMwNzE0MDAwMDAwWhcNMzQx
# MDEzMjM1OTU5WjBIMQswCQYDVQQGEwJVUzEXMBUGA1UEChMORGlnaUNlcnQsIElu
# Yy4xIDAeBgNVBAMTF0RpZ2lDZXJ0IFRpbWVzdGFtcCAyMDIzMIICIjANBgkqhkiG
# 9w0BAQEFAAOCAg8AMIICCgKCAgEAo1NFhx2DjlusPlSzI+DPn9fl0uddoQ4J3C9I
# o5d6OyqcZ9xiFVjBqZMRp82qsmrdECmKHmJjadNYnDVxvzqX65RQjxwg6seaOy+W
# ZuNp52n+W8PWKyAcwZeUtKVQgfLPywemMGjKg0La/H8JJJSkghraarrYO8pd3hkY
# hftF6g1hbJ3+cV7EBpo88MUueQ8bZlLjyNY+X9pD04T10Mf2SC1eRXWWdf7dEKEb
# g8G45lKVtUfXeCk5a+B4WZfjRCtK1ZXO7wgX6oJkTf8j48qG7rSkIWRw69XloNpj
# sy7pBe6q9iT1HbybHLK3X9/w7nZ9MZllR1WdSiQvrCuXvp/k/XtzPjLuUjT71Lvr
# 1KAsNJvj3m5kGQc3AZEPHLVRzapMZoOIaGK7vEEbeBlt5NkP4FhB+9ixLOFRr7St
# FQYU6mIIE9NpHnxkTZ0P387RXoyqq1AVybPKvNfEO2hEo6U7Qv1zfe7dCv95NBB+
# plwKWEwAPoVpdceDZNZ1zY8SdlalJPrXxGshuugfNJgvOuprAbD3+yqG7HtSOKmY
# CaFxsmxxrz64b5bV4RAT/mFHCoz+8LbH1cfebCTwv0KCyqBxPZySkwS0aXAnDU+3
# tTbRyV8IpHCj7ArxES5k4MsiK8rxKBMhSVF+BmbTO77665E42FEHypS34lCh8zrT
# ioPLQHsCAwEAAaOCAYswggGHMA4GA1UdDwEB/wQEAwIHgDAMBgNVHRMBAf8EAjAA
# MBYGA1UdJQEB/wQMMAoGCCsGAQUFBwMIMCAGA1UdIAQZMBcwCAYGZ4EMAQQCMAsG
# CWCGSAGG/WwHATAfBgNVHSMEGDAWgBS6FtltTYUvcyl2mi91jGogj57IbzAdBgNV
# HQ4EFgQUpbbvE+fvzdBkodVWqWUxo97V40kwWgYDVR0fBFMwUTBPoE2gS4ZJaHR0
# cDovL2NybDMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0VHJ1c3RlZEc0UlNBNDA5NlNI
# QTI1NlRpbWVTdGFtcGluZ0NBLmNybDCBkAYIKwYBBQUHAQEEgYMwgYAwJAYIKwYB
# BQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBYBggrBgEFBQcwAoZMaHR0
# cDovL2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0VHJ1c3RlZEc0UlNBNDA5
# NlNIQTI1NlRpbWVTdGFtcGluZ0NBLmNydDANBgkqhkiG9w0BAQsFAAOCAgEAgRrW
# 3qCptZgXvHCNT4o8aJzYJf/LLOTN6l0ikuyMIgKpuM+AqNnn48XtJoKKcS8Y3U62
# 3mzX4WCcK+3tPUiOuGu6fF29wmE3aEl3o+uQqhLXJ4Xzjh6S2sJAOJ9dyKAuJXgl
# nSoFeoQpmLZXeY/bJlYrsPOnvTcM2Jh2T1a5UsK2nTipgedtQVyMadG5K8TGe8+c
# +njikxp2oml101DkRBK+IA2eqUTQ+OVJdwhaIcW0z5iVGlS6ubzBaRm6zxbygzc0
# brBBJt3eWpdPM43UjXd9dUWhpVgmagNF3tlQtVCMr1a9TMXhRsUo063nQwBw3syY
# nhmJA+rUkTfvTVLzyWAhxFZH7doRS4wyw4jmWOK22z75X7BC1o/jF5HRqsBV44a/
# rCcsQdCaM0qoNtS5cpZ+l3k4SF/Kwtw9Mt911jZnWon49qfH5U81PAC9vpwqbHkB
# 3NpE5jreODsHXjlY9HxzMVWggBHLFAx+rrz+pOt5Zapo1iLKO+uagjVXKBbLafIy
# mrLS2Dq4sUaGa7oX/cR3bBVsrquvczroSUa31X/MtjjA2Owc9bahuEMs305MfR5o
# cMB3CtQC4Fxguyj/OOVSWtasFyIjTvTs0xf7UGv/B3cfcZdEQcm4RtNsMnxYL2dH
# ZeUbc7aZ+WssBkbvQR7w8F/g29mtkIBEr4AQQYowggauMIIElqADAgECAhAHNje3
# JFR82Ees/ShmKl5bMA0GCSqGSIb3DQEBCwUAMGIxCzAJBgNVBAYTAlVTMRUwEwYD
# VQQKEwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xITAf
# BgNVBAMTGERpZ2lDZXJ0IFRydXN0ZWQgUm9vdCBHNDAeFw0yMjAzMjMwMDAwMDBa
# Fw0zNzAzMjIyMzU5NTlaMGMxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2Vy
# dCwgSW5jLjE7MDkGA1UEAxMyRGlnaUNlcnQgVHJ1c3RlZCBHNCBSU0E0MDk2IFNI
# QTI1NiBUaW1lU3RhbXBpbmcgQ0EwggIiMA0GCSqGSIb3DQEBAQUAA4ICDwAwggIK
# AoICAQDGhjUGSbPBPXJJUVXHJQPE8pE3qZdRodbSg9GeTKJtoLDMg/la9hGhRBVC
# X6SI82j6ffOciQt/nR+eDzMfUBMLJnOWbfhXqAJ9/UO0hNoR8XOxs+4rgISKIhjf
# 69o9xBd/qxkrPkLcZ47qUT3w1lbU5ygt69OxtXXnHwZljZQp09nsad/ZkIdGAHvb
# REGJ3HxqV3rwN3mfXazL6IRktFLydkf3YYMZ3V+0VAshaG43IbtArF+y3kp9zvU5
# EmfvDqVjbOSmxR3NNg1c1eYbqMFkdECnwHLFuk4fsbVYTXn+149zk6wsOeKlSNbw
# sDETqVcplicu9Yemj052FVUmcJgmf6AaRyBD40NjgHt1biclkJg6OBGz9vae5jtb
# 7IHeIhTZgirHkr+g3uM+onP65x9abJTyUpURK1h0QCirc0PO30qhHGs4xSnzyqqW
# c0Jon7ZGs506o9UD4L/wojzKQtwYSH8UNM/STKvvmz3+DrhkKvp1KCRB7UK/BZxm
# SVJQ9FHzNklNiyDSLFc1eSuo80VgvCONWPfcYd6T/jnA+bIwpUzX6ZhKWD7TA4j+
# s4/TXkt2ElGTyYwMO1uKIqjBJgj5FBASA31fI7tk42PgpuE+9sJ0sj8eCXbsq11G
# deJgo1gJASgADoRU7s7pXcheMBK9Rp6103a50g5rmQzSM7TNsQIDAQABo4IBXTCC
# AVkwEgYDVR0TAQH/BAgwBgEB/wIBADAdBgNVHQ4EFgQUuhbZbU2FL3MpdpovdYxq
# II+eyG8wHwYDVR0jBBgwFoAU7NfjgtJxXWRM3y5nP+e6mK4cD08wDgYDVR0PAQH/
# BAQDAgGGMBMGA1UdJQQMMAoGCCsGAQUFBwMIMHcGCCsGAQUFBwEBBGswaTAkBggr
# BgEFBQcwAYYYaHR0cDovL29jc3AuZGlnaWNlcnQuY29tMEEGCCsGAQUFBzAChjVo
# dHRwOi8vY2FjZXJ0cy5kaWdpY2VydC5jb20vRGlnaUNlcnRUcnVzdGVkUm9vdEc0
# LmNydDBDBgNVHR8EPDA6MDigNqA0hjJodHRwOi8vY3JsMy5kaWdpY2VydC5jb20v
# RGlnaUNlcnRUcnVzdGVkUm9vdEc0LmNybDAgBgNVHSAEGTAXMAgGBmeBDAEEAjAL
# BglghkgBhv1sBwEwDQYJKoZIhvcNAQELBQADggIBAH1ZjsCTtm+YqUQiAX5m1tgh
# QuGwGC4QTRPPMFPOvxj7x1Bd4ksp+3CKDaopafxpwc8dB+k+YMjYC+VcW9dth/qE
# ICU0MWfNthKWb8RQTGIdDAiCqBa9qVbPFXONASIlzpVpP0d3+3J0FNf/q0+KLHqr
# hc1DX+1gtqpPkWaeLJ7giqzl/Yy8ZCaHbJK9nXzQcAp876i8dU+6WvepELJd6f8o
# VInw1YpxdmXazPByoyP6wCeCRK6ZJxurJB4mwbfeKuv2nrF5mYGjVoarCkXJ38SN
# oOeY+/umnXKvxMfBwWpx2cYTgAnEtp/Nh4cku0+jSbl3ZpHxcpzpSwJSpzd+k1Os
# Ox0ISQ+UzTl63f8lY5knLD0/a6fxZsNBzU+2QJshIUDQtxMkzdwdeDrknq3lNHGS
# 1yZr5Dhzq6YBT70/O3itTK37xJV77QpfMzmHQXh6OOmc4d0j/R0o08f56PGYX/sr
# 2H7yRp11LB4nLCbbbxV7HhmLNriT1ObyF5lZynDwN7+YAN8gFk8n+2BnFqFmut1V
# wDophrCYoCvtlUG3OtUVmDG0YgkPCr2B2RP+v6TR81fZvAT6gt4y3wSJ8ADNXcL5
# 0CN/AAvkdgIm2fBldkKmKYcJRyvmfxqkhQ/8mJb2VVQrH4D6wPIOK+XW+6kvRBVK
# 5xMOHds3OBqhK/bt1nz8MIIFjTCCBHWgAwIBAgIQDpsYjvnQLefv21DiCEAYWjAN
# BgkqhkiG9w0BAQwFADBlMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQg
# SW5jMRkwFwYDVQQLExB3d3cuZGlnaWNlcnQuY29tMSQwIgYDVQQDExtEaWdpQ2Vy
# dCBBc3N1cmVkIElEIFJvb3QgQ0EwHhcNMjIwODAxMDAwMDAwWhcNMzExMTA5MjM1
# OTU5WjBiMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYD
# VQQLExB3d3cuZGlnaWNlcnQuY29tMSEwHwYDVQQDExhEaWdpQ2VydCBUcnVzdGVk
# IFJvb3QgRzQwggIiMA0GCSqGSIb3DQEBAQUAA4ICDwAwggIKAoICAQC/5pBzaN67
# 5F1KPDAiMGkz7MKnJS7JIT3yithZwuEppz1Yq3aaza57G4QNxDAf8xukOBbrVsaX
# bR2rsnnyyhHS5F/WBTxSD1Ifxp4VpX6+n6lXFllVcq9ok3DCsrp1mWpzMpTREEQQ
# Lt+C8weE5nQ7bXHiLQwb7iDVySAdYyktzuxeTsiT+CFhmzTrBcZe7FsavOvJz82s
# NEBfsXpm7nfISKhmV1efVFiODCu3T6cw2Vbuyntd463JT17lNecxy9qTXtyOj4Da
# tpGYQJB5w3jHtrHEtWoYOAMQjdjUN6QuBX2I9YI+EJFwq1WCQTLX2wRzKm6RAXwh
# TNS8rhsDdV14Ztk6MUSaM0C/CNdaSaTC5qmgZ92kJ7yhTzm1EVgX9yRcRo9k98Fp
# iHaYdj1ZXUJ2h4mXaXpI8OCiEhtmmnTK3kse5w5jrubU75KSOp493ADkRSWJtppE
# GSt+wJS00mFt6zPZxd9LBADMfRyVw4/3IbKyEbe7f/LVjHAsQWCqsWMYRJUadmJ+
# 9oCw++hkpjPRiQfhvbfmQ6QYuKZ3AeEPlAwhHbJUKSWJbOUOUlFHdL4mrLZBdd56
# rF+NP8m800ERElvlEFDrMcXKchYiCd98THU/Y+whX8QgUWtvsauGi0/C1kVfnSD8
# oR7FwI+isX4KJpn15GkvmB0t9dmpsh3lGwIDAQABo4IBOjCCATYwDwYDVR0TAQH/
# BAUwAwEB/zAdBgNVHQ4EFgQU7NfjgtJxXWRM3y5nP+e6mK4cD08wHwYDVR0jBBgw
# FoAUReuir/SSy4IxLVGLp6chnfNtyA8wDgYDVR0PAQH/BAQDAgGGMHkGCCsGAQUF
# BwEBBG0wazAkBggrBgEFBQcwAYYYaHR0cDovL29jc3AuZGlnaWNlcnQuY29tMEMG
# CCsGAQUFBzAChjdodHRwOi8vY2FjZXJ0cy5kaWdpY2VydC5jb20vRGlnaUNlcnRB
# c3N1cmVkSURSb290Q0EuY3J0MEUGA1UdHwQ+MDwwOqA4oDaGNGh0dHA6Ly9jcmwz
# LmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydEFzc3VyZWRJRFJvb3RDQS5jcmwwEQYDVR0g
# BAowCDAGBgRVHSAAMA0GCSqGSIb3DQEBDAUAA4IBAQBwoL9DXFXnOF+go3QbPbYW
# 1/e/Vwe9mqyhhyzshV6pGrsi+IcaaVQi7aSId229GhT0E0p6Ly23OO/0/4C5+KH3
# 8nLeJLxSA8hO0Cre+i1Wz/n096wwepqLsl7Uz9FDRJtDIeuWcqFItJnLnU+nBgMT
# dydE1Od/6Fmo8L8vC6bp8jQ87PcDx4eo0kxAGTVGamlUsLihVo7spNU96LHc/RzY
# 9HdaXFSMb++hUD38dglohJ9vytsgjTVgHAIDyyCwrFigDkBjxZgiwbJZ9VVrzyer
# bHbObyMt9H5xaiNrIv8SuFQtJ37YOtnwtoeW/VvRXKwYw02fc7cBqZ9Xql4o4rmU
# MYIDdjCCA3ICAQEwdzBjMQswCQYDVQQGEwJVUzEXMBUGA1UEChMORGlnaUNlcnQs
# IEluYy4xOzA5BgNVBAMTMkRpZ2lDZXJ0IFRydXN0ZWQgRzQgUlNBNDA5NiBTSEEy
# NTYgVGltZVN0YW1waW5nIENBAhAFRK/zlJ0IOaa/2z9f5WEWMA0GCWCGSAFlAwQC
# AQUAoIHRMBoGCSqGSIb3DQEJAzENBgsqhkiG9w0BCRABBDAcBgkqhkiG9w0BCQUx
# DxcNMjQwMjI5MDEwNzEyWjArBgsqhkiG9w0BCRACDDEcMBowGDAWBBRm8CsywsLJ
# D4JdzqqKycZPGZzPQDAvBgkqhkiG9w0BCQQxIgQgp/K43eGtQsaJYsVrLj2nd6yb
# eZC4oM9qoBouY6ZEUXwwNwYLKoZIhvcNAQkQAi8xKDAmMCQwIgQg0vbkbe10IszR
# 1EBXaEE2b4KK2lWarjMWr00amtQMeCgwDQYJKoZIhvcNAQEBBQAEggIAbbaS5pM4
# myh4OqF3t8kiwH7o2wwfJxKBvBrPfbm1Tl/GqsDoj3N/3HmK6hBStduk1QSfLz6O
# 3ZukgOxkzmtsg/L6qFzcez+lguuord/Nx2khyvp0206H9DFf3Bsk3xh7oy54rlLK
# sxiw2ZHqRMAwU3f4aIZT+rtNmemMb6kf+lQzQH0yrh4F43eZ85gDnn1aXa2udsmy
# 3ChJivUelZNHke8sK68CMIncAo0yX1QUa7Jc1iP+QpSWm+KITcWpa4ZwDp1jHcn6
# hEkHi4RF36fzk00sFiF5V20Q2gYeJYfZSoJ/0dCA2R5GKbTAtunb2Mwb6dl/nQy7
# rwSpEicA653uYRO1/LHxkhfnnDbiKo5/vnmkLV9V1V6BHl8g9RWcXawoRfl6Whom
# JelM05NCJRFz30P2pgvrLA/WE4/7N8lfMPYo1HNoKpVnclP+gV1kb3U0utwwQHAT
# g5BtjYuui55v/yb9JWI3L5VzjTjXJhvlK4Wg9dG+V4ndf/d3G8+KpBzwBs+Q4y8g
# 1aPwPvhKmcPssFXAicHxNuInX84nZXO3revZDSJG/ZFwRBXx6Z1GF+R2UF4yRjF6
# LUq+MD9u8VFlJK1hd78xefF0noTNeaCYaEJO19fbnuNtbd54rgwhqe8V/txarhqA
# 9sQd0LCw7BFpORe1/bf7Ns/5QsZ1Ar0+Eyk=
# SIG # End signature block
