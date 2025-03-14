<# Path location of installation
D:\AppDynamics for application
D:\AppDynamics\backup for backups
D:\AppDynamics\binaries for installs

D:\AppDynamics[below agent type]
SmartAgent
DotNetAgent
DotNetAgentConfig
JavaAgent
MachineAgent
NetViz

#>

<#
    $configFile = 'config.xml',
    $instrumentedServices = @('W3SVC'),
    $agentServiceName = 'AppDynamics.Agent.Coordinator',
    $agentRegistryBranch = 'HKLM:\Software\AppDynamics\dotNet Agent',
    $agentPackageName = 'AppDynamics .NET Agent'

    $winstonConfigFile = 'winstonConfigFile.xml'
#>

param (
    $environment = 'prod',
    $controller = '',
    $controllerName = "$($controller).saas.appdynamics.com",
    $controllerUri = "https://$($controllerName)",

    $agentServiceName = 'AppDynamics.Agent.Coordinator_service',
    $instrumentedServices = @('W3SVC'),
    $agentRegistryBranch = 'HKLM:\Software\AppDynamics\dotNet Agent',
    $agentPackageName = 'AppDynamics .NET Agent',
    $agentType = 'DotNetAgent',
    $DotNetAgentConfig = 'DotNetAgentConfig',
    $agentBase = 'dotNetAgentSetup*64',
    $configDir = 'Config',
    $configFile = 'config.xml',

    $AppDynamicsAgentsParentPath = 'D:\AppDynamics',
    $AppDynamicsAgentPath = "$($AppDynamicsAgentsParentPath)\$($agentType)",
    $AppDynamicsAgentConfigPath = "$($AppDynamicsAgentsParentPath)\$($DotNetAgentConfig)",
    $DotNetAgentFolder = "$($AppDynamicsAgentConfigPath)\", # Will be updated if found in registry.
    $originalDotNetAgentFolder = '',
    $targetDotNetAgentFolder = $DotNetAgentFolder, # Used to verify if old is the same as new directory.
    $AppDynamicsAgentsMsiFilesPath = "$($AppDynamicsAgentsParentPath)\binaries",
    $winstonConfigFile = "$($AppDynamicsAgentsMsiFilesPath)\winstonConfigFile.xml",
    $msiExecCommandFile = 'dotNetAgentInstall',
    $msiExecCommandFileExt = 'cmd',
    $AppDynamicsBackupPath = "$($AppDynamicsAgentsParentPath)\backup",
    $AppDynamicsAgentBackupPath = "$($AppDynamicsBackupPath)\$($agentType)",
    # .NET Agent only needs configuration files backed up. Exes should not be.
    $AppDynamicsAgentConfigBackupPath = "$($AppDynamicsBackupPath)\$($DotNetAgentConfig)",
    $logFileDirectory = "$($AppDynamicsAgentConfigBackupPath)\Logs",
    $priorFolders = @('Config', 'Data'), # Data directory contains configuration files as well
    $AppDynamicsAgent = "$($AppDynamicsAgentsParentPath)\$($agentType)",
    $installDir = "$($AppDynamicsAgentsParentPath)\$($agentType)",
    $AppDAgentPathBaseFullName = $installDir, # Updated if service is installed

    $agentFolderFound = $false,
    [ValidateSet('Move', 'Remove', 'Keep', 'NotFound')]
    [string]$priorFolderOperation # Old Configuration will remain as is if neither value set.
)

<#

#>
. D:\AppDynamics\binaries\AppDAgentCommon.ps1

# Reset all installation operations in case you are doing multiple rounds of debugging and testing.
# Otherwise, variables remain in memory in VS Code
$install = $uninstall = $upgrade = ''
$today = Get-Date
$timestamp = $today.ToString('yyyyMMddHHmm')

# exit code must be an integer so any items with substeps need exitcode to be parent step
$stepExitCode = $stepNumber = 1
$stepDescription = "Agent Installation Script Started at $($timestamp)."
Write-Output "Step:$($stepNumber)"
Write-Output $stepDescription

$stepExitCode = $stepNumber = 2
$stepDescription = "Checking that backup folder $($AppDynamicsAgentBackupPath) exists."
Write-Output "Step:$($stepNumber)"
Write-Output $stepDescription

try
{
    New-AppDFolder -ParentPath $AppDynamicsAgentConfigBackupPath -Path $priorFolders -exitCode $stepExitCode
    # Need Logs directory created for silent install to work correctly. Parent of logfile must exist.
    New-AppDFolder -ParentPath $AppDynamicsAgentConfigBackupPath -Path 'Logs' -exitCode $stepExitCode
}
catch
{
    Write-Output 'Unable to create backup directories.'
    $Error[0]
    Exit $stepExitCode
}

$stepExitCode = $stepNumber = 3
$stepDescription = 'Checking that installation file exists.'
Write-Output "Step:$($stepNumber)"
Write-Output $stepDescription

$stepNumber = 3.1
$stepDescription = "Test-Path for binaries path $($AppDynamicsAgentsMsiFilesPath)."
Write-Output "Step:$($stepNumber)"
Write-Output $stepDescription

if (!(Test-Path $AppDynamicsAgentsMsiFilesPath))
{
    Write-Output "$($AppDynamicsAgentsMsiFilesPath) directory Not Found."
    $Error[0]
    Exit $stepExitCode
}

$stepNumber = 3.2
$agentMsi = "$($agentBase)*.msi"
$stepDescription = "Test-Path for $($agentMsi)."
Write-Output "Step:$($stepNumber)"
Write-Output $stepDescription

<#
Append wildcard to handle any version of agent zip
Get-ChildItem on wildcard might return 0 results, so Test-Path required.
#>

$downloadedAgentMsi = (Get-ChildItem -Path $AppDynamicsAgentsMsiFilesPath -Filter $agentMsi).FullName

if (!(Test-Path $downloadedAgentMsi))
{
    Write-Output "$($downloadedAgentMsi) Not Found."
    $Error[0]
    Exit $stepExitCode
}

$stepExitCode = $stepNumber = 4
$stepDescription = 'Checking for current agent existence: service or just directory.'
Write-Output "Step:$($stepNumber)"
Write-Output $stepDescription

try
{
    $stepNumber = '4.1'
    $stepDescription = 'Checking for service.'
    Write-Output "Step:$($stepNumber)"
    Write-Output $stepDescription
    
    $agentService = Get-Service -ServiceName $agentServiceName -ErrorAction Stop
    $agentServiceName = $agentService.Name
    $agentService = Get-CimInstance -class Win32_Service `
    | Where-Object { $_.Name -eq $agentServiceName }
    $servicePathName = ($agentService | Select-Object PathName).PathName 
    if ($servicePathName)
    {
        $exePathAgentService = ($servicePathName).Trim('"')
        $currentInstallDirectory = (Get-Item $exePathAgentService).DirectoryName

        $stepNumber = '4.1.1'
        $stepDescription = "Agent Installed as a service. Path to exe: $($servicePathName)."
        Write-Output "Step:$($stepNumber)"
        Write-Output $stepDescription
        Write-Output "Current Installation Directory: $($currentInstallDirectory)"

        if ($currentInstallDirectory -eq $AppDynamicsAgentPath)
        {
            $upgrade = $true
        }
        else
        {
            $uninstall = $true
        }
    }   
}
catch
{
    $stepNumber = '4.1.2'
    $stepDescription = 'Agent not installed as a service.'
    Write-Output "Step:$($stepNumber)"
    Write-Output $stepDescription
    $install = $true
}

$stepNumber = '4.2'
$stepDescription = 'Querying registry for agent configuration directory.'
Write-Output "Step:$($stepNumber)"
Write-Output $stepDescription

if (Test-Path $agentRegistryBranch)
{
    if ($registryDotNetAgent = Get-ItemProperty -Path $agentRegistryBranch -ErrorAction SilentlyContinue)
    {
        $DotNetAgentFolder = $registryDotNetAgent.DotNetAgentFolder
        $originalDotNetAgentFolder = $DotNetAgentFolder
        if ($originalDotNetAgentFolder -eq $targetDotNetAgentFolder)
        {
            $stepNumber = '4.2.1'
            $stepDescription = 'Registry key for agent configuration directory matches desired directory.'
            Write-Output "Step:$($stepNumber)"
            Write-Output $stepDescription
            $priorFolderOperation = 'Keep'
        }
        else
        {
            $stepNumber = '4.2.2'
            $stepDescription = 'Registry key for agent configuration directory does not match desired directory.'
            Write-Output "Step:$($stepNumber)"
            Write-Output $stepDescription
        }
    }
    else
    {
        $stepNumber = '4.2.3'
        $stepDescription = 'Registry key for agent configuration directory does not exist.'
        Write-Output "Step:$($stepNumber)"
        Write-Output $stepDescription
    }
}
else
{
    $stepNumber = '4.2.3'
    $stepDescription = 'Registry branch for agent configuration directory does not exist.'
    Write-Output "Step:$($stepNumber)"
    Write-Output $stepDescription
}

$stepNumber = '4.3'
$stepDescription = 'Checking for existing agent configuration directory'
Write-Output "Step:$($stepNumber)"
Write-Output $stepDescription

$potentialParentFolder = 'ProgramData\AppDynamics\DotNetAgent'

$potentialParentDrive = 'D:'
$potentialConfigFolder = "$($potentialParentDrive)\$($potentialParentFolder)"

$defaultParentDrive = 'C:'
$defaultConfigFolder = "$($defaultParentDrive)\$($potentialParentFolder)"

$possibleFolders = @($DotNetAgentFolder, $AppDynamicsAgentConfigPath, $AppDynamicsAgentConfigBackupPath, $potentialConfigFolder, $defaultConfigFolder)
$configFileProps = Get-AppDConfigFileBackupFolder -ParentPath $possibleFolders -childPath $configDir -childFileName $configFile -exitCode $stepExitCode
$DotNetAgentFolder = $configFileProps.confParentPath

if ($DotNetAgentFolder)
{
    $stepNumber = '4.4'
    $stepDescription = "Configuration Folder value set: $($DotNetAgentFolder)"
    Write-Output "Step:$($stepNumber)"
    Write-Output $stepDescription
    try
    {
        if (Test-Path $DotNetAgentFolder)
        {
            $stepNumber = '4.4.1'
            $stepDescription = "Agent configuration directory found: $($DotNetAgentFolder)"
            Write-Output "Step:$($stepNumber)"
            Write-Output $stepDescription
            $originalDotNetAgentFolder = $DotNetAgentFolder
            $agentFolderFound = $true
            $configFileFullName = [IO.Path]::Combine($DotNetAgentFolder, $configDir, $configFile)
            if (Test-Path $configFileFullName)
            {
                $stepNumber = '4.4.1.1'
                $stepDescription = "Configuration file $($configFileFullName) found."
                Write-Output "Step:$($stepNumber)"
                Write-Output $stepDescription
                $agentFolderFound = $true
            }
            else 
            {
                $stepNumber = '4.4.1.2'
                $stepDescription = "Configuration file $($configFileFullName) not found."
                Write-Output "Step:$($stepNumber)"
                Write-Output $stepDescription
            }
        }
        else
        {
            $stepNumber = '4.4.2'
            $stepDescription = "Agent configuration directory $($DotNetAgentFolder) not found."
            Write-Output "Step:$($stepNumber)"
            Write-Output $stepDescription
        }
    }
    catch
    {
        Write-Output "Configuration Folder value is $($DotNetAgentFolder) but does not exist."
        $Error[0]
        Exit $stepExitCode
    }
}
else
{
    $stepNumber = '4.5'
    $stepDescription = 'Agent configuration directory not found.'
    Write-Output "Step:$($stepNumber)"
    Write-Output $stepDescription
}

$stepExitCode = $stepNumber = 5
$stepDescription = 'Checking for folders to backup.'
Write-Output "Step:$($stepNumber)"
Write-Output $stepDescription

if ($agentFolderFound)
{
    if ($DotNetAgentFolder -eq $AppDynamicsAgentConfigBackupPath)
    {
        $stepDescription = 'Current Configuration Folder is the same as backup, so no need to copy.'
        $stepNumber = '5.1'
        Write-Output "Step:$($stepNumber)"
        Write-Output $stepDescription
    }
    elseif ($priorFolders -and $DotNetAgentFolder)
    {
        try
        {
            $stepDescription = "Copying select folders from prior agent to backups:`n$($priorFolders -join "`n")"
            $stepNumber = '5.2'
            Write-Output "Step:$($stepNumber)"
            Write-Output $stepDescription
            Copy-AppDBackupFolders -ParentPath $DotNetAgentFolder -Path $priorFolders -Destination $AppDynamicsAgentConfigBackupPath -exitCode $stepExitCode                
        }
        catch
        {
            Write-Output "Unable to copy select folders from prior agent to backups:`n$($priorFolders)"
            $Error[0]
            Exit $stepExitCode
        }
    }
    else
    {
        $stepDescription = 'No prior agent configuration directory to backup.'
        $stepNumber = '5.3'
        Write-Output "Step:$($stepNumber)"
        Write-Output $stepDescription
    }
}
else
{
    $stepDescription = 'No prior agent configuration directory to backup.'
    $stepNumber = '5.4'
    Write-Output "Step:$($stepNumber)"
    Write-Output $stepDescription
}

$stepDescription = 'Installation Decision Tree: New Install, Upgrade, or Uninstall wrong directory and install to new directory.'
$stepNumber = '6'
Write-Output "Step:$($stepNumber)"
Write-Output $stepDescription

if ($uninstall)
{
    $stepExitCode = $stepNumber = 7
    $stepDescription = 'Uninstall.'
    Write-Output "Step:$($stepNumber)"
    Write-Output $stepDescription

    $stepDescription = "Uninstalling $($agentPackageName) msi package."
    $stepNumber = '7.1'
    Write-Output "Step:$($stepNumber)"
    Write-Output $stepDescription

    # Replace with msiexec code as *-Package not available everywhere
    <#
    
        if (Get-Package -Name $agentPackageName -ProviderName msi)
    {
        $global:ProgressPreference = 'SilentlyContinue'

        try
        {
            Uninstall-Package -Name $agentPackageName -ProviderName msi -ErrorAction Stop
            
            Remove-Item -Path $agentRegistryBranch -Recurse    
        }
        catch
        {
            Write-Output "Unable to uninstall $($agentPackageName)."
            $Error[0]
            Exit $stepExitCode
        }
    }    
    else
    {
        Write-Output "MSI $($agentPackageName) not found."
        $Error[0]
        Exit $stepExitCode
    }

    #>
    try
    {

        $agentMsiInfo = Get-ItemProperty 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*' | Where-Object { $_.DisplayName -eq "$agentPackageName" }
        $agentMsiGUID = $agentMsiInfo.PSChildName


        #msiexec /x {10CA4E7D-55D0-4D5F-B4FF-7E33A57850FA} /q /norestart /lv D:\AppDynamics\backup\DotNetAgentConfig\Logs\uninstall.log
        $msiExecCommandFile = 'dotNetAgentUninstall'
        $msiExecCommand = @"
msiexec /x $($agentMsiGUID) /q /norestart /lv $($logFileDirectory)\AgentUninstall.log
"@
        $msiExecCommandFileName = "$($msiExecCommandFile).$($msiExecCommandFileExt)"
        $msiExecCommandFileFullName = Join-Path $AppDynamicsAgentsMsiFilesPath $msiExecCommandFileName
        $msiExecCommand | Set-Content $msiExecCommandFileFullName
        Get-Content $msiExecCommandFileFullName

        $stepDescription = 'Run uninstall script.'
        $stepNumber = '7.2'
        Write-Output "Step:$($stepNumber)"
        Write-Output $stepDescription

        New-AppDProcessWrapper -AppDAgentPathBaseFullName $AppDynamicsAgentsMsiFilesPath -targetScriptBaseName $msiExecCommandFile -targetScriptExtension $msiExecCommandFileExt -wrapperString '' -silentSuffix '' -exitCode $stepExitCode

        # Configuration Data Path left in registry even after uninstall.
        Remove-Item -Path $agentRegistryBranch -Recurse
        
    }
    catch
    {
        Write-Output "Unable to uninstall $($agentPackageName)."
        $Error[0]
        Exit $stepExitCode
    }

    $stepDescription = 'Proceed to Installation.'
    $stepNumber = '7.2'
    Write-Output "Step:$($stepNumber)"
    Write-Output $stepDescription

    $install = $true
}

if ($upgrade)
{
    $priorFolderOperation = 'Keep'
    
    $stepExitCode = $stepNumber = 8
    $stepDescription = 'Upgrade.'
    Write-Output "Step:$($stepNumber)"
    Write-Output $stepDescription

    $stepDescription = 'Create upgrade script.'
    $stepNumber = '8.1'
    Write-Output "Step:$($stepNumber)"
    Write-Output $stepDescription
    $msiExecCommandFile = 'dotNetAgentUpgrade'
    $msiExecCommand = @"
msiexec /i $($downloadedAgentMsi) /q /lv $($logFileDirectory)\AgentUpgrade.log INSTALLDIR=$($installDir)
"@
    $msiExecCommandFileName = "$($msiExecCommandFile).$($msiExecCommandFileExt)"
    $msiExecCommandFileFullName = Join-Path $AppDynamicsAgentsMsiFilesPath $msiExecCommandFileName
    $msiExecCommand | Set-Content $msiExecCommandFileFullName
    Get-Content $msiExecCommandFileFullName

    $stepDescription = 'Run upgrade script.'
    $stepNumber = '8.2'
    Write-Output "Step:$($stepNumber)"
    Write-Output $stepDescription

    New-AppDProcessWrapper -AppDAgentPathBaseFullName $AppDynamicsAgentsMsiFilesPath -targetScriptBaseName $msiExecCommandFile -targetScriptExtension $msiExecCommandFileExt -wrapperString '' -silentSuffix '' -exitCode $stepExitCode
}

elseif ($install)
{
    $stepExitCode = $stepNumber = 9
    $stepDescription = 'Install.'
    Write-Output "Step:$($stepNumber)"
    Write-Output $stepDescription

    $stepDescription = 'Create install script.'
    $stepNumber = '9.1'
    Write-Output "Step:$($stepNumber)"
    Write-Output $stepDescription

    $msiExecCommandFile = 'dotNetAgentInstall'
    $msiExecCommand = @"
msiexec /i $($downloadedAgentMsi) /q /norestart /lv $($logFileDirectory)\AgentInstaller.log AD_SetupFile=$($winstonConfigFile) INSTALLDIR=$($installDir) DOTNETAGENTFOLDER=$($AppDynamicsAgentConfigPath)
"@
    $msiExecCommandFileName = "$($msiExecCommandFile).$($msiExecCommandFileExt)"
    $msiExecCommandFileFullName = Join-Path $AppDynamicsAgentsMsiFilesPath $msiExecCommandFileName
    $msiExecCommand | Set-Content $msiExecCommandFileFullName
    Get-Content $msiExecCommandFileFullName

    $stepDescription = 'Run install script.'
    $stepNumber = '9.2'
    Write-Output "Step:$($stepNumber)"
    Write-Output $stepDescription

    New-AppDProcessWrapper -AppDAgentPathBaseFullName $AppDynamicsAgentsMsiFilesPath -targetScriptBaseName $msiExecCommandFile -targetScriptExtension $msiExecCommandFileExt -wrapperString '' -silentSuffix '' -exitCode $stepExitCode
}

$stepDescription = 'Check to see if old configuration directory should be kept, moved, or removed.'
$stepNumber = '10'
Write-Output "Step:$($stepNumber)"
Write-Output $stepDescription

if (!($priorFolderOperation))
{
    $priorFolderOperation = 'Keep'
}
Write-Output "Prior Folder Operation is $($priorFolderOperation)."

if ($originalDotNetAgentFolder)
{
    try 
    {
        Test-Path $originalDotNetAgentFolder
        try
        {
            $stepDescription = 'Restore backups to new installation.'
            $stepNumber = '10.1'
            Write-Output "Step:$($stepNumber)"
            Write-Output $stepDescription

            if ($upgrade)
            {
                $stepDescription = 'Inplace upgrade so backups are the same as current. No need to restore.'
                $stepNumber = '10.1.1'
                Write-Output "Step:$($stepNumber)"
                Write-Output $stepDescription
            }
            elseif (!($agentFolderFound))
            {
                $stepDescription = 'No prior agent configuration directory to restore.'
                $stepNumber = '10.1.2'
                Write-Output "Step:$($stepNumber)"
                Write-Output $stepDescription
            }
            else
            {
                $stepDescription = 'Restoring backups.'
                $stepNumber = '10.1.3'
                Write-Output "Step:$($stepNumber)"
                Write-Output $stepDescription
                Copy-AppDBackupFolders -ParentPath $AppDynamicsAgentConfigBackupPath -Path $priorFolders -Destination $AppDynamicsAgentConfigPath -exitCode $stepExitCode -desiredErrorAction 'Stop'
            }            
        }
        catch
        {
            Write-Output 'Unable to copy backups to new installation.'
            $Error[0]
            Exit $stepExitCode
        }
        if ($DotNetAgentFolder -eq $AppDynamicsAgentConfigBackupPath)
        {
            $priorFolderOperation = 'Keep'
        }
        if ($priorFolderOperation -eq 'Keep') 
        {
            $stepDescription = "Prior Folder Operation is $($priorFolderOperation). No action taken."
            $stepNumber = '10.2'
            Write-Output "Step:$($stepNumber)"
            Write-Output $stepDescription
        }
        elseif ($priorFolderOperation -eq 'Move') 
        {
            $stepDescription = "Prior Folder Operation is $($priorFolderOperation)."
            $stepNumber = '10.3'
            Write-Output "Step:$($stepNumber)"
            Write-Output $stepDescription
            $originalDotNetAgentFolder = Get-Item $originalDotNetAgentFolder
            $originalDotNetAgentFolderName = $originalDotNetAgentFolder.Name
            $originalDotNetAgentFolderParent = $originalDotNetAgentFolder.Parent
            $priorAgentBackupPath = Join-Path $originalDotNetAgentFolderParent.FullName "$($originalDotNetAgentFolderName)_$($timestamp)"
            $stepDescription = "Backing up with Move-Item $($originalDotNetAgentFolder) to $($priorAgentBackupPath)."
            $stepNumber = '10.3.1'
            Write-Output "Step:$($stepNumber)"
            Write-Output $stepDescription

            try
            {
                Move-Item -Path $($originalDotNetAgentFolder) -Destination $($priorAgentBackupPath) -Force -ErrorAction Stop
            }
            catch
            {
                Write-Output "Unable to move $($originalDotNetAgentFolder) to $($priorAgentBackupPath)."
                $Error[0]
                Exit $stepExitCode
            }
        }
        elseif ($priorFolderOperation -eq 'Remove') 
        {
            $stepDescription = "Prior Folder Operation is $($priorFolderOperation)."
            $stepNumber = '10.4'
            Write-Output "Step:$($stepNumber)"
            Write-Output $stepDescription

            try
            {
                Remove-Item -Path $($originalDotNetAgentFolder) -Recurse -Force -ErrorAction Stop
            }
            catch
            {
                Write-Output "Unable to delete $($originalDotNetAgentFolder)."
                $Error[0]
                Exit $stepExitCode
            }
        }        
    }
    catch 
    {
        $stepDescription = 'No prior folder was found.'
        $stepNumber = '10.1.1'
        Write-Output "Step:$($stepNumber)"
        Write-Output $stepDescription
    }    
}
else 
{
    $stepDescription = 'No prior folder was found.'
    $stepNumber = '10.5'
    Write-Output "Step:$($stepNumber)"
    Write-Output $stepDescription
}

$stepNumber = - 1
$stepDescription = 'Agent Installation Script END.'
Write-Output "Step:$($stepNumber)"
Write-Output $stepDescription
