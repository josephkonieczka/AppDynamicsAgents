param (
    $environment = 'prod',

    $agentServiceName = 'NetVizService',
    $agentType = 'NetViz',
 
    $agentBase = 'appd-netviz-x64-windows',
    $AppDynamicsAgentsParentPath = 'D:\AppDynamics',

    $AppDynamicsAgentsZipFilesPath = "$($AppDynamicsAgentsParentPath)\binaries",
    $AppDynamicsAgent = "$($AppDynamicsAgentsParentPath)\$($agentType)",
    $unzippedAgentPath = "$($AppDynamicsAgentsParentPath)\$($agentType)",
    $AppDAgentPathBaseFullName = $unzippedAgentPath, # Updated if service is installed

    $agentFolderFound = $false,
    $secondsFileLockCheck = 60,
    $oldValue = "set quiet=NO`r`nset check_admin=YES",
    $newValue = "set quiet=YES`r`nset check_admin=NO",
    [ValidateSet('Move', 'Remove')]
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
$stepDescription = 'Checking that installation zip file exists.'
Write-Output "Step:$($stepNumber)"
Write-Output $stepDescription

$stepNumber = 2.1
$stepDescription = "Test-Path for binaries path $($AppDynamicsAgentsZipFilesPath)."
Write-Output "Step:$($stepNumber)"
Write-Output $stepDescription

if (!(Test-Path $AppDynamicsAgentsZipFilesPath))
{
    Write-Output "$($AppDynamicsAgentsZipFilesPath) directory Not Found."
    $Error[0]
    Exit $stepExitCode
}

$stepNumber = 2.2
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

$stepExitCode = $stepNumber = 3
$stepDescription = 'Checking for current agent existence: service or just directory.'
Write-Output "Step:$($stepNumber)"
Write-Output $stepDescription

try
{
    $AppDAgentPathBaseFullName = Get-AppDAgentPathBaseFullName -agentServiceName $agentServiceName
    $stepDescription = "Agent Installed as a service. Path to exe: $($AppDAgentPathBaseFullName)."
    $stepNumber = 3.1
    Write-Output "Step:$($stepNumber)"
    Write-Output $stepDescription
    $agentFolderFound = $true
    try
    {
        $stepDescription = 'Uninstalling agent service.'
        $stepNumber = 3.1.1
        Write-Output "Step:$($stepNumber)"
        Write-Output $stepDescription
        New-AppDProcessWrapper -AppDAgentPathBaseFullName $AppDAgentPathBaseFullName -targetScriptBaseName 'uninstall' -targetScriptExtension 'bat' -wrapperString '' -oldValue $oldValue -newValue $newValue -exitCode $stepExitCode
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
    $stepNumber = 3.2
    Write-Output "Step:$($stepNumber)"
    Write-Output $stepDescription

    if (Test-Path $unzippedAgentPath)
    {
        $agentFolderFound = $true
        $stepDescription = "Correct agent directory $($unzippedAgentPath) exists."
        $stepNumber = '3.2.1'
        Write-Output "Step:$($stepNumber)"
        Write-Output $stepDescription
    }
}

if ($priorFolderOperation)
{
    $stepNumber = '4'
    $stepDescription = "Prior Folder Handling Operation is $($priorFolderOperation)."
    Write-Output "Step:$($stepNumber)"
    Write-Output $stepDescription
}
    
# Remove the prior directory and all its contents recursively
if ($priorFolderOperation -eq 'Remove') 
{
    $stepDescription = 'Deleting prior agent directory.'
    $stepNumber = '4.1'
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
    $stepNumber = '4.2'
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

try
{
    $stepDescription = 'Installing agent service.'
    $stepNumber = 6
    Write-Output "Step:$($stepNumber)"
    Write-Output $stepDescription
    New-AppDProcessWrapper -AppDAgentPathBaseFullName $unzippedAgentPath -targetScriptBaseName 'install' -targetScriptExtension 'bat' -wrapperString '' -oldValue $oldValue -newValue $newValue -exitCode $stepExitCode

    Set-Service -Name $agentServiceName -StartupType Automatic
    Start-Service -ServiceName $agentServiceName

}
catch
{
    Write-Output 'Unable to install agent service.'
    $Error[0]
    Exit $stepExitCode
} 



$stepNumber = - 1
$stepDescription = 'Agent Installation Script END.'
Write-Output "Step:$($stepNumber)"
Write-Output $stepDescription
