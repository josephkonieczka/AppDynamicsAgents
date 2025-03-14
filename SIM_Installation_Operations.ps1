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
Add option for additional command line startup options
-Dappdynamics.http.proxyHost
-Dappdynamics.http.proxyPort
#>

param (
    $environment = 'prod',
    $controller = '',
    $controllerName = "$($controller).saas.appdynamics.com",
    $controllerUri = "https://$($controllerName)",

    $agentServiceName = 'Appdynamics Machine Agent',
    $agentType = 'MachineAgent',
    $agentBase = 'machineagent-bundle-64bit-windows',
    $AppDynamicsAgentsParentPath = 'D:\AppDynamics',
    $AppDynamicsAgent = "$($AppDynamicsAgentsParentPath)\$($agentType)",
    $AppDynamicsAgentConfigFolder = "$($AppDynamicsAgent)\conf",
    $AppDynamicsAgentsZipFilesPath = "$($AppDynamicsAgentsParentPath)\binaries",
    $AppDynamicsBackupPath = "$($AppDynamicsAgentsParentPath)\backup",
    $AppDynamicsAgentBackupPath = "$($AppDynamicsBackupPath)\$($agentType)",
    $confDir = 'conf',
    $confBackup = "$($AppDynamicsAgentBackupPath)\$($confDir)",
    $configFile = 'controller-info.xml',
    $controllerInfoXMLBackup = "$($confBackup)\$($configFile)",
    $monitorsBackupPath = "$($AppDynamicsAgentBackupPath)\monitors",
    $controllerInfoXML = "$($AppDynamicsAgentConfigFolder)\controller-info.xml",
    $unzippedAgentPath = "$($AppDynamicsAgentsParentPath)\$($agentType)",
    $AppDAgentPathBaseFullName = $unzippedAgentPath, # Updated if service is installed
    $potentialPriorPath = 'D:\Program Files\AppDynamics\MachineAgent',
    $priorFolders = @('conf', 'local-scripts'),
    $customExtensions = @(''), #ProcessMonitor, ServerTools
    $sigarFilesCopy = $false,
    $agentFolderFound = $false,
    $secondsFileLockCheck = 60,
    <#
    JVM Options such as proxy config vars
    #>
    $JVMOptions = '-Dappdynamics.http.proxyHost=DMZHOST -Dappdynamics.http.proxyPort=9092',
    #$JVMOptions = '-Dmetric.http.listener=false -Dmetric.http.listener.port=8095',
    $matchComputername = '*FW*',
    $oldValue = 'WScript.StdIn.Read\(1\)',
    $newValue = '',
    [ValidateSet('Move', 'Remove', 'Overwrite', 'Keep')]
    [string]$priorFolderOperation # Unzip will overwrite if parameter not set
)

. D:\AppDynamics\binaries\AppDAgentCommon.ps1

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
    New-AppDFolder -ParentPath $AppDynamicsAgentBackupPath -Path $priorFolders -exitCode $stepExitCode
    # Added check for custom extensions as it would error out if it path is empty.
    if ($customExtensions)
    {
        New-AppDFolder -ParentPath $monitorsBackupPath -Path $customExtensions -exitCode $stepExitCode 
    }    
}
catch
{
    Write-Output 'Unable to create backup directories.'
    $Error[0]
    Exit $stepExitCode
}

$stepExitCode = $stepNumber = 3
$stepDescription = 'Checking that installation zip file exists.'
Write-Output "Step:$($stepNumber)"
Write-Output $stepDescription

$stepNumber = 3.1
$stepDescription = "Test-Path for binaries path $($AppDynamicsAgentsZipFilesPath)."
Write-Output "Step:$($stepNumber)"
Write-Output $stepDescription

if (!(Test-Path $AppDynamicsAgentsZipFilesPath))
{
    Write-Output "$($AppDynamicsAgentsZipFilesPath) directory Not Found."
    $Error[0]
    Exit $stepExitCode
}

$stepNumber = 3.2
$stepDescription = "Test-Path for $($agentBase)*.zip."
Write-Output "Step:$($stepNumber)"
Write-Output $stepDescription

<#
Append wildcard to handle any version of agent zip
Get-ChildItem on wildcard might return 0 results, so Test-Path required.
#>

$agentZip = "$($agentBase)*.zip"
$downloadedAgentZip = (Get-ChildItem -Path $AppDynamicsAgentsZipFilesPath -Filter $agentZip).FullName

if (!(Test-Path $downloadedAgentZip))
{
    Write-Output "$($downloadedAgentZip) Not Found."
    $Error[0]
    Exit $stepExitCode
}

$stepExitCode = $stepNumber = 4
$stepDescription = 'Checking for current agent existence: service or just directory.'
Write-Output "Step:$($stepNumber)"
Write-Output $stepDescription

try
{
    $AppDAgentPathBaseFullName = Get-AppDAgentPathBaseFullName -agentServiceName $agentServiceName
    $stepDescription = "Agent Installed as a service. Path to exe: $($AppDAgentPathBaseFullName)."
    $stepNumber = 4.1
    Write-Output "Step:$($stepNumber)"
    Write-Output $stepDescription
    $agentFolderFound = $true
    try
    {
        $stepDescription = 'Uninstalling agent service.'
        $stepNumber = 4.1.1
        Write-Output "Step:$($stepNumber)"
        Write-Output $stepDescription
        New-AppDProcessWrapper -AppDAgentPathBaseFullName $AppDAgentPathBaseFullName -targetScriptBaseName 'uninstallservice' -oldValue $oldValue -newValue $newValue -exitCode $stepExitCode
    }
    catch
    {
        Write-Output 'Unable to uninstall agent service.'
        $Error[0]
        Exit $stepExitCode
    } 
}
catch 
{
    $stepDescription = 'Agent Not Installed as a service.'
    $stepNumber = 4.2
    Write-Output "Step:$($stepNumber)"
    Write-Output $stepDescription
}

# Common block if service or directory - check process existence and backup current folders
# Hardware Monitoring process stays running for a short period of time, so waiting until it is no longer running
$stepDescription = 'Check for running Hardware Monitoring process.'
$stepNumber = '4.2.2'
Write-Output "Step:$($stepNumber)"
Write-Output $stepDescription

Test-AppDWindowsProcess

$stepDescription = 'Check for prior agent configuration directory to backup.'
$stepNumber = '4.3'
Write-Output "Step:$($stepNumber)"
Write-Output $stepDescription

if ($AppDAgentPathBaseFullName) { $possibleFolders = @($AppDAgentPathBaseFullName) }
$possibleFolders += @($unzippedAgentPath, $AppDynamicsAgentBackupPath)
if ($potentialPriorPath) { $possibleFolders += $potentialPriorPath }

$configFileProps = Get-AppDConfigFileBackupFolder -ParentPath $possibleFolders -childPath $confDir -childFileName $configFile -exitCode $stepExitCode
$AppDAgentPathBaseFullName = $configFileProps.confParentPath

$stepDescription = 'Creating backups.'
$stepNumber = '4.3.1'
Write-Output "Step:$($stepNumber)"
Write-Output $stepDescription

if ($AppDAgentPathBaseFullName -eq $AppDynamicsAgentBackupPath) 
{
    $stepDescription = 'Backups same as config source folder. Nothing to do.'
    $stepNumber = '4.3.2'
    Write-Output "Step:$($stepNumber)"
    Write-Output $stepDescription

    $priorFolderOperation = 'Keep'
}
else
{

    if ($priorFolders )
    {
        try
        {
            $stepDescription = "Copying select folders from prior agent to backups:`n$($priorFolders)."
            $stepNumber = '4.3.2'
            Write-Output "Step:$($stepNumber)"
            Write-Output $stepDescription
 
            Copy-AppDBackupFolders -ParentPath $AppDAgentPathBaseFullName -Path $priorFolders -Destination $AppDynamicsAgentBackupPath -exitCode $stepExitCode
        }
        catch
        {
            Write-Output "Unable to copy select folders from prior agent to backups:`n$($priorFolders)"
            $Error[0]
            Exit $stepExitCode
        } 
    }

    # Custom Extensions are under \monitors
    $monitorsPath = Join-Path $AppDAgentPathBaseFullName 'monitors'
    $monitorsBackupPath = Join-Path $AppDynamicsAgentBackupPath 'monitors'

    if ($customExtensions)
    {
        try
        {
            $stepDescription = "Copying Custom Machine Agent Extension: $($customExtensions)."
            $stepNumber = '4.3.3'
            Write-Output "Step:$($stepNumber)"
            Write-Output $stepDescription
            # Path to custom extensions is actually under \monitors
            # Check to see if the extension folder already exists. It might not.
            $extensionCount = 0
            foreach ($extension in $customExtensions)
            {
                $extensionCount++
                $extensionPath = Join-Path $monitorsPath $extension
                $extensionPath
                $extensionBackupPath = Join-Path $monitorsBackupPath $extension
                $extensionBackupPath
                if (Test-Path $extensionPath)
                {
                    try
                    {
                        $stepDescription = "Copying Custom Machine Agent Extension $($extension)."
                        $stepNumber = "4.3.2.$($extensionCount)"
                        Write-Output "Step:$($stepNumber)"
                        Write-Output $stepDescription
                        Copy-AppDBackupFolders -ParentPath $monitorsPath -Path $extension -Destination $monitorsBackupPath -exitCode $stepExitCode
                        Copy-Item $extensionPath $monitorsBackupPath -Force -Recurse #-WhatIf
                    }
                    catch
                    {
                        Write-Output "Unable to backup Custom Machine Agent Extension $($extension)"
                        $Error[0]
                        Exit $stepExitCode
                    }
                }
                else
                {
                    Write-Output "Custom Machine Agent Extension $($extension) Not Found"
                }
            }
        }
        catch
        {
            Write-Output "Unable to copy Custom Extensions $($customExtensions)"
            $Error[0]
            Exit $stepExitCode
        }
    }

    # Machine Agent configuration settings like Server Visibility under \extensions
    $machineAgentConfiguratonExtensionsPath = Join-Path $AppDAgentPathBaseFullName 'extensions'
    $machineAgentConfiguratonExtensionsBackupPath = Join-Path $AppDynamicsAgentBackupPath 'extensions'

    $configurationFilesCount = 0
    $stepDescription = 'Copying Machine Agent configuration files.'
    $stepNumber = '4.3.4'
    Write-Output "Step:$($stepNumber)"
    Write-Output $stepDescription

    # Unknown number of configuration directories as new agent versions sometimes add one or more
    # Create backup directories and then add yml files to correct folder
    $machineAgentConfigurationExtensions = (Get-ChildItem $machineAgentConfiguratonExtensionsPath -Directory)
    $machineAgentConfigurationExtensionsNames = ($machineAgentConfigurationExtensions).Name
    New-AppDFolder -ParentPath $machineAgentConfiguratonExtensionsBackupPath -Path $machineAgentConfigurationExtensionsNames -exitCode $stepExitCode
    foreach ($configurationExtensionFolder in $machineAgentConfigurationExtensions)
    {
        $configurationExtensionFolderName = $configurationExtensionFolder.Name
        $configurationExtensionFolderFullName = $configurationExtensionFolder.FullName
        $configurationExtensionFolderBackupPath = Join-Path $machineAgentConfiguratonExtensionsBackupPath $configurationExtensionFolderName
        New-AppDFolder -ParentPath $configurationExtensionFolderBackupPath -Path 'conf' -exitCode $stepExitCode
        $configurationFilesCount++
        $stepDescription = "Copying Machine Agent configuration files from  $($configurationExtensionFolderFullName) to $($configurationExtensionFolderBackupPath)."
        $stepNumber = "4.3.4.$($configurationFilesCount)"
        Write-Output "Step:$($stepNumber)"
        Write-Output $stepDescription
        
        $ymlConfFolder = Join-Path $configurationExtensionFolderFullName 'conf'
        if (Test-Path $ymlConfFolder)
        {
            $ymlDestinationFolder = Join-Path $configurationExtensionFolderBackupPath 'conf'
            try
            {
                Copy-AppDBackupFolders -ParentPath $configurationExtensionFolderFullName -Path 'conf' -Destination $configurationExtensionFolderBackupPath -exitCode $stepExitCode
            }
            catch
            {
                Write-Output "Unable to copy $($ymlConfFolder) to $($ymlDestinationFolder)."
                $Error[0]
                Exit $stepExitCode
            }
        } 
    }
}

# Cautious approach to retain prior backup if necessary
if (!($priorFolderOperation))
{
    $priorFolderOperation = 'Keep'
}

$stepNumber = '4.4'
$stepDescription = "Prior Folder Handling Operation is $($priorFolderOperation)."
Write-Output "Step:$($stepNumber)"
Write-Output $stepDescription

# Remove the prior directory and all its contents recursively
if ($priorFolderOperation -eq 'Remove') 
{
    $stepDescription = 'Deleting prior agent directory.'
    $stepNumber = '4.4.1'
    Write-Output "Step:$($stepNumber)"
    Write-Output $stepDescription
    if (Test-Path $AppDAgentPathBaseFullName)
    {
        Write-Output "Deleting $($agentType) directory: $($AppDAgentPathBaseFullName)"
        try
        {
            Remove-Item -Path $AppDAgentPathBaseFullName -Recurse -Force -ErrorAction Stop
            Write-Output "$($agentType) directory deleted: $($AppDAgentPathBaseFullName)" 
        }
        catch
        {
            Write-Output "Unable to delete $($AppDAgentPathBaseFullName)\$($agentType)"
            $Error[0]
            Exit $stepExitCode
        }
    }
    else
    {
        Write-Output "$($agentType) directory not found: $AppDAgentPathBaseFullName"
    }
}
# Move folder and append timestamp
elseif ($priorFolderOperation -eq 'Move') 
{	
    $priorAgentBackupPath = "$($AppDAgentPathBaseFullName)_$($timestamp)"
    $stepDescription = "Backing up with Move-Item $($AppDAgentPathBaseFullName) to $($priorAgentBackupPath)."
    $stepNumber = '4.4.2'
    Write-Output "Step:$($stepNumber)"
    Write-Output $stepDescription
    if (Test-Path $AppDAgentPathBaseFullName)
    {
        Write-Output "Moving $($agentType) directory: $($AppDAgentPathBaseFullName)"
        try
        {
            Move-Item -Path $AppDAgentPathBaseFullName -Destination $priorAgentBackupPath -Force -ErrorAction Stop
            Write-Output "$($agentType) directory moved: $($AppDAgentPathBaseFullName)" 
        }
        catch
        {
            Write-Output "Unable to move $($AppDAgentPathBaseFullName)\$($agentType)"
            $Error[0]
            Exit $stepExitCode
        }
    }
    else
    {
        Write-Output "$($agentType) directory not found: $AppDAgentPathBaseFullName"
    }
}


$stepExitCode = $stepNumber = 5
$stepDescription = 'Unzipping agent.'
Write-Output "Step:$($stepNumber)"
Write-Output $stepDescription

#Suppress Progress Bar ...
$global:ProgressPreference = 'SilentlyContinue'

try
{
    Expand-Archive -LiteralPath $downloadedAgentZip -DestinationPath $unzippedAgentPath -Force
}
catch
{
    Write-Output "Unable to unzip $($downloadedAgentZip) to $($AppDynamicsAgentsParentPath)"
    $Error[0]
    Exit $stepExitCode
}

$stepExitCode = $stepNumber = 6
$stepDescription = 'Configuration settings and backup restoration.'
Write-Output "Step:$($stepNumber)"
Write-Output $stepDescription

if (Test-Path $AppDynamicsAgentBackupPath)
{
    #Just copy from agent to agent at root
    $stepDescription = "Restoring folders from $($AppDynamicsAgentBackupPath) $($unzippedAgentPath)."
    $stepNumber = '6.1'
    Write-Output "Step:$($stepNumber)"
    Write-Output $stepDescription
    try
    {
        $restoreFolders = (Get-ChildItem $AppDynamicsAgentBackupPath -Directory)
        Copy-AppDBackupFolders -ParentPath $AppDynamicsAgentBackupPath -Path $restoreFolders -Destination $unzippedAgentPath -exitCode $stepExitCode
    }
    catch
    {
        Write-Output 'Unable to restore backup folders.'
        $Error[0]
        Exit $stepExitCode
    }
}
else
{
    $stepNumber = '6.2'
    $stepDescription = 'No backups found. No restoration of prior configuration settings.'
    Write-Output "Step:$($stepNumber)"
    Write-Output $stepDescription
}

$stepExitCode = $stepNumber = 7
$stepDescription = 'Installing agent as a service.'
Write-Output "Step:$($stepNumber)"
Write-Output $stepDescription
try
{
    if ($env:COMPUTERNAME -like $matchComputername)
    {
        $additionalSuffix = $JVMOptions
    }
    New-AppDProcessWrapper -AppDAgentPathBaseFullName $unzippedAgentPath -targetScriptBaseName 'installservice' -oldValue $oldValue -newValue $newValue -exitCode $stepExitCode -additionalSuffix $additionalSuffix
}
catch
{
    Write-Output 'Unable to install agent as a service'
    $Error[0]
    Exit $stepExitCode
}

$stepExitCode = $stepNumber = 8
$stepDescription = 'Check agent details - version # and controller connection.'
Write-Output "Step:$($stepNumber)"
Write-Output $stepDescription

Write-Output "Sleeping for $($secondsFileLockCheck) seconds to allow agent to start and log connection details."
Start-Sleep $secondsFileLockCheck

$machineAgentLogFolderPath = Join-Path $unzippedAgentPath 'logs'
$machineAgentLogPath = Join-Path $machineAgentLogFolderPath 'machine-agent.log'

if (!(Test-Path $machineAgentLogFolderPath))
{
    Write-Output "$($machineAgentLogFolderPath) directory not created. Check controller-info.xml."
    $Error[0]
    Exit $stepExitCode
}
elseif (!(Test-Path $machineAgentLogPath))
{
    Write-Output "$($machineAgentLogPath) file not created. Check controller-info.xml."
    $Error[0]
    Exit $stepExitCode
}
else
{
    $loggedAgentVersion = Select-String -Pattern 'Agent Version' -Path $machineAgentLogPath | Select-Object -Last 1
    $loggedControllerConnection = Select-String -Pattern "XML Controller Info Resolver found controller host \[$($controllerName)]" -Path $machineAgentLogPath | Select-Object -Last 1

    if ($loggedAgentVersion -and $loggedControllerConnection)
    {
        Write-Output 'Logged Agent Version:'
        Write-Output "$($loggedAgentVersion)"
        Write-Output 'Logged Controller Connection:'
        Write-Output "$($loggedControllerConnection)"
    }
    else
    {
        Write-Output "Machine agent was unable to successfully connect to $($controllerName)."
        $Error[0]
        Exit $stepExitCode
    }

}

$stepNumber = - 1
$stepDescription = 'Agent Installation Script END.'
Write-Output "Step:$($stepNumber)"
Write-Output $stepDescription
