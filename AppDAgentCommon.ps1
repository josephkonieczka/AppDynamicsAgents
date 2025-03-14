function Get-AppDAgentPathBaseFullName
{
    param (
        [Parameter(Mandatory = $True)]
        [ValidateScript( {
                if (Get-Service $_)
                {
                    $True
                }
                else
                {
                    Throw "$_ is an invalid Service Name"
                }
            })] $agentServiceName   
    )
    $agentService = Get-CimInstance -class Win32_Service | Where-Object { $_.Name -eq $agentServiceName }
    $servicePathName = ($agentService | Select-Object PathName).PathName
    $exePathAgentService = ($servicePathName).Trim('"')
    $agentPathDirs = ($exePathAgentService -split '\\')
    $agentPathDrive = $agentPathDirs[0]
    $agentBasePathFullName = $agentPathDrive
    $agentBasePathLastDirIndex = $agentPathDirs.Length - 3
    $agentBasePathDirs = $agentPathDirs[1 .. $agentBasePathLastDirIndex]
    $agentBasePathDirs.ForEach({ $agentBasePathFullName = Join-Path $agentBasePathFullName $_ })
    Write-Output $agentBasePathFullName    
}


<#
NetViz
$oldValue = 'quiet=NO',
$newValue = 'quiet=YES',

Windows Machine Agent
$oldValue = 'WScript.StdIn.Read\(1\)',
$newValue = '',

$targetScriptBaseName = 'installservice',
$targetScriptBaseName = 'uninstallservice',

Proxy config vars to add to command line cscript
-Dappdynamics.http.proxyHost=HOSTNAME -Dappdynamics.http.proxyPort=PORTNUM
#>
function New-AppDProcessWrapper
{
    param (
        [parameter(Mandatory = $true)]
        $AppDAgentPathBaseFullName,
        [parameter(Mandatory = $true)]
        $targetScriptBaseName,
        $oldValue,
        $newValue,
        $wrapperString = 'cscript.exe',
        $targetScriptExtension = 'vbs',
        $targetScriptWrapperExtension = 'cmd',
        $silentSuffix = '_silent',
        $additionalSuffix = '',
        [parameter(Mandatory = $true)]
        $exitCode
    )
    
    $targetScriptBaseName = Join-Path $AppDAgentPathBaseFullName $targetScriptBaseName
    $targetScript = "$($targetScriptBaseName).$($targetScriptExtension)"

    $silentTargetScript = "$($targetScriptBaseName)$($silentSuffix).$($targetScriptExtension)"
    try 
    {
        Test-Path $targetScript | Out-Null
        if ($oldValue)
        {
            $rawTexttargetScript = Get-Content $targetScript -Raw
            #$rawTexttargetScript -replace 'WScript.StdIn.Read\(1\)', '' `
            $rawTexttargetScript -replace $oldValue, $newValue `
            | Set-Content $silentTargetScript
        }
        if ($wrapperString)
        {
            $silentTargetWrapper = "$($targetScriptBaseName)_silent.$($targetScriptWrapperExtension)"
            if ($additionalSuffix)
            {
                # JVM Options are on separate lines
                "$($wrapperString) `"$($silentTargetScript)`" $($additionalSuffix)"
                "$($wrapperString) `"$($silentTargetScript)`" $($additionalSuffix)"| Set-Content $silentTargetWrapper
            }
            else
            {
                "$($wrapperString) `"$($silentTargetScript)`"" | Set-Content $silentTargetWrapper
            }
            
        }
        else
        {
            $silentTargetWrapper = $silentTargetScript
        }
        
        $psi = New-Object System.Diagnostics.ProcessStartInfo
        $psi.FileName = $silentTargetWrapper
        $psi.Verb = 'runas'
        $psi.RedirectStandardError = $true
        $psi.RedirectStandardOutput = $true
        $psi.CreateNoWindow = $true
        $psi.UseShellExecute = $false

        try
        {
            $targetProcess = New-Object System.Diagnostics.Process
            $targetProcess.StartInfo = $psi
            $targetProcess.Start() | Out-Null
            $stdout = $targetProcess.StandardOutput.ReadToEnd()
            $stderr = $targetProcess.StandardError.ReadToEnd()
            $targetProcess.WaitForExit()
            $exitCode = $targetProcess.ExitCode
            Write-Output "Silent Wrapper $($silentTargetWrapper) for $($targetScript) script started successfully with administrator privileges."
            Write-Output "`nOutput and Exit Code from $($targetScript)"
            Write-Output "STDOUT: $stdout"
            Write-Output "STDERR: $stderr"
            Write-Output "Exit Code: $($exitCode)"
            if ($exitCode -ne 0)
            {
                $Error[0]
                Exit $exitCode
            }
        }
        catch
        {
            Write-Output "`nFailed to start the target script with administrator privileges: $_"
            $Error[0]
            Exit $exitCode
        }
    }
    catch
    {
        Write-Output "target script file not found: $targetScript"
        Exit $exitCode
    }
}


function Test-AppDWindowsProcess
{
    param (
        $processName = 'cscript.exe',
        $commandLine = 'windows-stat.vbs',
        $processDescription = 'HardwareMonitor windows-stat',
        $secondsFileLockCheck = 60
    )
    #::TODO:: Add foreach logic to handle multiple processes matching criteria
    #Limiting to first process found for this iteration of function. 
    $processInfo = (Get-CimInstance -Query "SELECT * from Win32_Process WHERE Name LIKE '%$processName%' and CommandLine LIKE '%$commandLine%'")
    $processInfo = $processInfo | Select-Object -First 1
    $processPid = ($processInfo).ProcessId

    if ($processPid)
    {
        Write-Output "$($processDescription) PID is $($processPid)"
        while (Get-Process -Id $processPid -ErrorAction SilentlyContinue)
        {
            Write-Output "Looping via sleeping for $($secondsFileLockCheck) seconds to wait for $($processDescription) to end."
            Start-Sleep $secondsFileLockCheck
        }
    }
    else
    {
        Write-Output "$($processDescription) PID is not running."
    }
}


function New-AppDFolder
{
    param (
        [Parameter(Mandatory = $true)]
        $ParentPath,
        [parameter(Mandatory = $true)]
        [string[]]$Path,
        [parameter(Mandatory = $true)]
        $exitCode
    )
    foreach ($folder in $Path)
    {
        $folderFullName = Join-Path $ParentPath $folder
        if (!(Test-Path $folderFullName))
        {
            try
            {
                Write-Output "Creating $($folderFullName) directory."
                New-Item -Path $folderFullName -ItemType Directory | Out-Null
            }
            catch
            {
                Write-Output "$($folderFullName) directory cannot be created."
                $Error[0]
                Exit $exitCode        
            }
        }
    }
}


function Copy-AppDBackupFolders
{
    param (
        [Parameter(Mandatory = $true)]
        [ValidateScript( {
                if (Test-Path $_)
                {
                    $True
                }
                else
                {
                    Throw "Cannot validate path $_"
                }
            })] $ParentPath,
        [parameter(Mandatory = $true)]
        [string[]]$Path,
        [parameter(Mandatory = $true)]
        $Destination,
        [parameter(Mandatory = $true)]
        $exitCode,
        $desiredErrorAction = 'SilentlyContinue'
    )
    foreach ($folder in $Path)
    {
        $sourceFolderFullName = Join-Path $ParentPath $folder
        try
        {
            Write-Output "Copying from $($sourceFolderFullName) to $($Destination)"
            Copy-Item "$($sourceFolderFullName)" $Destination -Recurse -Force -ErrorAction $desiredErrorAction #-WhatIf
        }
        catch
        {
            Write-Output "Unable to copy from $($sourceFolderFullName) to $($Destination)"
            $Error[0]
            Exit $exitCode
        }
    }    
}


function Get-AppDConfigFileBackupFolder
{
    param (
        [parameter(Mandatory = $true)]
        [string[]]$ParentPath,
        $childPath = 'conf',
        $childFileName = 'controller-info.xml',
        [parameter(Mandatory = $true)]
        $exitCode
    )
    foreach ($folder in $ParentPath)
    {
        $fileParentPath = Join-Path $folder $childPath
        $fileFullName = Join-Path $fileParentPath $childFileName
        if ((Test-Path $fileFullName))
        {
            $props = @{
                confParentPath     = $folder
                confPath           = $fileParentPath
                configFileFullName = $fileFullName
            }
            $AppDConfigFileBackupFolder = New-Object -TypeName PSObject -Property $props
            Write-Output $AppDConfigFileBackupFolder
            break       
        }
    }
}
