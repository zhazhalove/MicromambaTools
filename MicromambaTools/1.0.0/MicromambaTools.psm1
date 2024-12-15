
<#
.SYNOPSIS
Sets the MAMBA_ROOT_PREFIX environment variable for micromamba.

.DESCRIPTION
The `Initialize-MambaRootPrefix` function sets the environment variable `MAMBA_ROOT_PREFIX` to the specified path for micromamba. If no path is specified, it defaults to the `$env:APPDATA\micromamba` directory. This environment variable is typically used by micromamba to define the root directory for its environments.

.PARAMETER MAMBA_ROOT_PREFIX
The path to set for the `MAMBA_ROOT_PREFIX` environment variable. If not provided, the function defaults to `$env:APPDATA\micromamba`.

.EXAMPLE
Initialize-MambaRootPrefix -MAMBA_ROOT_PREFIX "C:\mamba"
This will set the `MAMBA_ROOT_PREFIX` environment variable to "C:\mamba".

.EXAMPLE
Initialize-MambaRootPrefix
This will set the `MAMBA_ROOT_PREFIX` environment variable to the default value of `$env:APPDATA\micromamba`.

.NOTES
https://mamba.readthedocs.io/en/latest/user_guide/concepts.html#root-prefix
#>
function Initialize-MambaRootPrefix {
    param (
        [Parameter(Mandatory = $false)]
        [string]$MAMBA_ROOT_PREFIX = "$env:APPDATA\micromamba"
    )
    # Set MAMBA_ROOT_PREFIX environment variable
    $env:MAMBA_ROOT_PREFIX = $MAMBA_ROOT_PREFIX
}


<#
.SYNOPSIS
    Checks if the specified micromamba environment exists.

.PARAMETER EnvName
    The name of the micromamba environment to check.

.RETURNS
    [bool] indicating if the environment exists.

.EXAMPLE
    Test-MicromambaEnvironment -EnvName "langchain"
#>
function Test-MicromambaEnvironment {
    param (
        [Parameter(Mandatory = $true)]
        [string]$EnvName
    )
    $pattern = "^\s*$EnvName\s*"
    $envList = & "$PSScriptRoot\micromamba.exe" env list | Select-String -Pattern $pattern
    return ($null -ne $envList -and $envList.Matches.Success -and $envList.Matches.Groups[0].Value.Trim() -eq $EnvName)
}

<#
.SYNOPSIS
    Installs a list of packages in a micromamba environment.

.PARAMETER EnvName
    The name of the micromamba environment where the packages will be installed.

.PARAMETER Packages
    A list of package names to be installed.

.PARAMETER TrustedHost
    A switch parameter to trust the certificate for pip installation. 
    If provided, adds the --trusted-host flag to the pip install command for trusted hosts (pypi.org and files.pythonhosted.org).
    https://pip.pypa.io/en/stable/topics/https-certificates/#

.EXAMPLE
    Install-PackagesInMicromambaEnvironment -EnvName "langchain" -Packages @("numpy", "pandas", "matplotlib")

.EXAMPLE
Install-PackagesInMicromambaEnvironment -EnvName "langchain" -Packages @("numpy", "pandas") -TrustedHost
#>
function Install-PackagesInMicromambaEnvironment {
    param (
        [Parameter(Mandatory = $true)]
        [string]$EnvName,

        [Parameter(Mandatory = $true)]
        [string[]]$Packages,

        [Parameter()]
        [switch]$TrustedHost
    )

    # Initialize an array to store results
    [hashtable[]]$results = @()

    # Write-Host "Installing packages in micromamba environment: $EnvName" -ForegroundColor Yellow

    foreach ($package in $Packages) {
        # Write-Host "Installing package: $package" -ForegroundColor Green

        try {
            if ($TrustedHost) {
                # Run the install command with trusted hosts
                & "$PSScriptRoot\micromamba.exe" run -n $EnvName pip install $package --trusted-host pypi.org --trusted-host files.pythonhosted.org *>&1 | Out-Null
            }
            else {
                # Run the install command without trusted hosts
                & "$PSScriptRoot\micromamba.exe" run -n $EnvName pip install $package *>&1 | Out-Null
            }

            # Check if the installation was successful
            if ($LASTEXITCODE -eq 0) {
                # Add success result for this package
                $results += @{
                    PackageName = $package
                    Success = $true
                }
                # Write-Host "Successfully installed package: $package" -ForegroundColor Cyan
            }
            else {
                # Add failure result for this package
                $results += @{
                    PackageName = $package
                    Success = $false
                }
                # Write-Host "Failed to install package: $package" -ForegroundColor Red
            }
        }
        catch {
            # In case of an exception, log the failure for this package
            $results += @{
                PackageName = $package
                Success = $false
            }
            # Write-Host "Error installing package: $package" -ForegroundColor Red
        }
    }

    # Write-Host "Package installation complete for environment: $EnvName" -ForegroundColor Yellow

    return $results
}



<#
.SYNOPSIS
    Creates the micromamba environment.

.PARAMETER EnvName
    The name of the micromamba environment to create. Required

.PARAMETER PythonVersion
    The Python version to use for the environment. Defaults to 3.11

.PARAMETER TrustedHost
    A switch parameter to trust the certificate for pip installation. 
    If provided, adds the --trusted-host flag to the pip install command for trusted hosts (pypi.org and files.pythonhosted.org).
    https://pip.pypa.io/en/stable/topics/https-certificates/#

.EXAMPLE
    Initialize-MicromambaEnvironment -EnvName "langchain"

.EXAMPLE
    Initialize-MicromambaEnvironment -EnvName "langchain" -PythonVersion "3.11" -TrustedHost
#>
function New-MicromambaEnvironment {
    param (
        [Parameter(Mandatory = $true)]
        [string]$EnvName,

        [Parameter(Mandatory = $false)]
        [string]$PythonVersion = '3.11',

        [Parameter()]
        [switch]$TrustedHost
    )

    # Write-Host "Creating micromamba environment: $EnvName" -ForegroundColor Yellow

    if ($TrustedHost) {
        # redirect output since uipath captures the output
        & "$PSScriptRoot\micromamba.exe" create -n $EnvName --yes --ssl-verify False python=$PythonVersion pip -c conda-forge *>&1 | Out-Null
    }
    else {
        # redirect output since uipath captures the output
        & "$PSScriptRoot\micromamba.exe" create -n $EnvName --yes python=$PythonVersion pip -c conda-forge *>&1 | Out-Null
    }
 
    if ($LASTEXITCODE -eq 0) {
        # Write-Host "Created the micromamba environment: $EnvName" -ForegroundColor Green
        $true
    }
    else {
        # Write-Host "Failed to create the micromamba environment: $EnvName" -ForegroundColor Red
        $false
    }
}


<#
.SYNOPSIS
    Executes a Python script within a specified virtual environment using micromamba.

.DESCRIPTION
    The `Invoke-PythonScript` function is designed to run a Python script located at a specified path
    within a virtual environment created using micromamba. Optional arguments can be passed to the 
    Python script as a string array. The output is expected to be in JSON format, which will be 
    converted into a PowerShell object.

.PARAMETER ScriptPath
    The full path to the Python script to execute. This parameter is mandatory.

.PARAMETER EnvName
    The name of the micromamba environment to use for executing the Python script. This parameter is mandatory.

.PARAMETER Arguments
    An optional array of strings representing the arguments to pass to the Python script. 
    If not provided, the script will be executed without additional arguments.

.OUTPUTS
    Returns a PowerShell object derived from the JSON output of the Python script.

.NOTES
    This function uses the `micromamba` tool to activate the specified environment and execute the Python script.
    The function assumes that the Python script outputs valid JSON.

    Author: [Your Name]
    Created: [Date]
    Version: 1.0

.EXAMPLE
    Example 1: Execute a Python script without additional arguments.
    
    $scriptPath = "C:\Scripts\example_script.py"
    $envName = "example_env"
    
    $result = Invoke-PythonScript -ScriptPath $scriptPath -EnvName $envName
    
    Write-Output $result

.EXAMPLE
    Example 2: Execute a Python script with additional arguments.
    
    $scriptPath = "C:\Scripts\example_script.py"
    $envName = "example_env"
    $arguments = @("--option1", "value1", "--option2", "value2")
    
    $result = Invoke-PythonScript -ScriptPath $scriptPath -EnvName $envName -Arguments $arguments
    
    Write-Output $result

.EXAMPLE
    Example 3: Handle errors when executing a Python script.
    
    $scriptPath = "C:\Scripts\example_script.py"
    $envName = "example_env"
    
    $result = Invoke-PythonScript -ScriptPath $scriptPath -EnvName $envName
    
    if ($null -eq $result) {
        Write-Host "The script failed to execute." -ForegroundColor Red
    } else {
        Write-Output $result
    }
#>
function Invoke-PythonScript {
    param (
        [Parameter(Mandatory = $true)]
        [string]$ScriptPath,

        [Parameter(Mandatory = $true)]
        [string]$EnvName,

        [Parameter(Mandatory = $false)]
        [string[]]$Arguments = @() # Optional parameter with a default value
    )

    Write-Host "Executing Python script: $ScriptPath" -ForegroundColor Yellow

    try {
        # Construct the argument string by joining all arguments with spaces
        $ArgumentString = $Arguments -join ' '

        # Execute the Python script with the constructed argument string
        $finalResult = & "$PSScriptRoot\micromamba.exe" run -n $EnvName python $ScriptPath $ArgumentString *>&1 | ConvertFrom-Json

        return $finalResult
    } catch {
        Write-Host "Error running Python script: $_" -ForegroundColor Red
        return $null
    }
}


<#
.SYNOPSIS
    Downloads and extracts the micromamba binary to the script's root directory and verifies its checksum.

.DESCRIPTION
    The `Get-MicromambaBinary` function downloads the micromamba binary and its SHA256 checksum file for Windows (64-bit) from the specified source URLs and extracts it into the directory where the script is located (`$PSScriptRoot`).
    It verifies the integrity of the downloaded binary by comparing its checksum with the provided SHA256 file.
    If the checksum verification fails, all downloaded and extracted files are removed.

.PARAMETER Url
    An optional URL to download the micromamba binary. Defaults to:
    "https://github.com/mamba-org/micromamba-releases/releases/latest/download/micromamba-win-64"

.PARAMETER ChecksumUrl
    An optional URL to download the SHA256 checksum file. Defaults to:
    "https://github.com/mamba-org/micromamba-releases/releases/latest/download/micromamba-win-64.sha256"

.EXAMPLE
    Get-MicromambaBinary
    Downloads, verifies, and extracts the micromamba binary from the default URLs.

.EXAMPLE
    Get-MicromambaBinary -Url "https://example.com/custom/micromamba.exe" -ChecksumUrl "https://example.com/custom/micromamba.sha256"
    Downloads, verifies, and extracts the micromamba binary from custom URLs.

.NOTES
    Ensure that the `tar` command is available in the environment if you use extracted tarballs in future versions.
#>
function Get-MicromambaBinary {
    param (
        [string]$Url = "https://github.com/mamba-org/micromamba-releases/releases/latest/download/micromamba-win-64",
        [string]$ChecksumUrl = "https://github.com/mamba-org/micromamba-releases/releases/latest/download/micromamba-win-64.sha256"
    )

    $destinationPath = $PSScriptRoot
    $binaryPath = Join-Path -Path $destinationPath -ChildPath "micromamba.exe"
    $checksumPath = Join-Path -Path $destinationPath -ChildPath "micromamba-win-64.sha256"

    try {
        # Download the micromamba binary
        Invoke-WebRequest -Uri $Url -OutFile $binaryPath -UseDefaultCredentials

        # Download the SHA256 checksum file
        Invoke-WebRequest -Uri $ChecksumUrl -OutFile $checksumPath -UseDefaultCredentials

        # Verify both files exist
        if (-not (Test-Path -Path $binaryPath) -or -not (Test-Path -Path $checksumPath)) {
            Remove-Item -Path $binaryPath, $checksumPath -ErrorAction SilentlyContinue
            return $false
        }

        # Read the checksum from the SHA256 file
        $expectedChecksum = (Get-Content -Path $checksumPath -ErrorAction Stop).Trim()

        # Compute the actual checksum of the downloaded binary
        $actualChecksum = (Get-FileHash -Path $binaryPath -Algorithm SHA256).Hash

        # Compare checksums
        if ($expectedChecksum -ne $actualChecksum) {
            # Remove files if checksum verification fails
            Remove-Item -Path $binaryPath, $checksumPath -ErrorAction SilentlyContinue
            return $false
        }

        # Return success if everything checks out
        return $true
    } catch {
        # Clean up files on error
        Remove-Item -Path $binaryPath, $checksumPath -ErrorAction SilentlyContinue
        return $false
    }
    finally {
        Remove-Item -Path $checksumPath -ErrorAction SilentlyContinue
    }
}


<#
.SYNOPSIS
    Cleans up a micromamba environment and unused cached packages.

.DESCRIPTION
    The `Remove-MicromambaEnvironment` function removes a specified micromamba environment using the `micromamba` tool.
    Additionally, it clears the cached packages using the `micromamba clean --all` command.

.PARAMETER EnvName
    The name of the micromamba environment to remove.

.RETURNS
    [bool] indicating if the cleanup was successful.

.EXAMPLE
    Cleanup-MicromambaEnvironment -EnvName "langchain"
#>
function Remove-MicromambaEnvironment {
    param (
        [Parameter(Mandatory = $true)]
        [string]$EnvName
    )

    try {
        # Remove the specified micromamba environment
        & "$PSScriptRoot\micromamba.exe" env remove -n $EnvName --yes *>&1 | Out-Null
        
        if ($LASTEXITCODE -ne 0) {
            return $false
        }

        # Clean up cached packages
        & "$PSScriptRoot\micromamba.exe" clean --all --yes *>&1 | Out-Null
        
        if ($LASTEXITCODE -ne 0) {
            return $false
        }

        return $true

    } catch {
        return $false
    }
}


<#
.SYNOPSIS
    Removes the micromamba executable, its root pefix directory, and unsets the MAMBA_ROOT_PREFIX environment variable.

.DESCRIPTION
    The `Remove-Micromamba` function deletes the `micromamba.exe` binary, the root prefix directory used by micromamba,
    and clears the `MAMBA_ROOT_PREFIX` environment variable.

.RETURNS
    [bool] indicating if the removal was successful.

.EXAMPLE
    Remove-Micromamba
#>
function Remove-Micromamba {
    try {
        # Define path for micromamba executable
        $micromambaPath = Join-Path -Path $PSScriptRoot -ChildPath "micromamba.exe"
        # Define path for micromamba root directory
        $micromambaRoot = $env:MAMBA_ROOT_PREFIX

        # Remove micromamba executable
        if (Test-Path -Path $micromambaPath) {
            Remove-Item -Path $micromambaPath -Force -Recurse
        }

        # Remove micromamba root directory
        if (Test-Path -Path $micromambaRoot) {
            Remove-Item -Path $micromambaRoot -Force -Recurse
        }

        # Unset the MAMBA_ROOT_PREFIX environment variable
        Remove-Item -Path Env:\MAMBA_ROOT_PREFIX -ErrorAction SilentlyContinue

        return $true
    } catch {
        return $false
    }
}


Export-ModuleMember -Function Remove-Micromamba, Remove-MicromambaEnvironment Get-MicromambaBinary, Invoke-PythonScript, New-MicromambaEnvironment, Install-PackagesInMicromambaEnvironment, Test-MicromambaEnvironment, Initialize-MambaRootPrefix