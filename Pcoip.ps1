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
    Set-EnvironmentVariableForScope -Name $EnvVariableName -Value $downloadDirectory -ScriptName $scriptName -Scope Process
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
