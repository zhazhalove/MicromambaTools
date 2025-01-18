
<#
.SYNOPSIS
    Sets the MAMBA_ROOT_PREFIX environment variable for micromamba.

.DESCRIPTION
    The `Initialize-MambaRootPrefix` function sets the environment variable `MAMBA_ROOT_PREFIX` to the specified path for micromamba. If no path is specified, it defaults to the `$env:APPDATA\micromamba` directory. This environment variable is typically used by micromamba to define the root directory for its environments.

.PARAMETER MAMBA_ROOT_PREFIX
    The path to set for the `MAMBA_ROOT_PREFIX` environment variable. If not provided, the function defaults to `$env:APPDATA\micromamba`.

.OUTPUTS
    $null

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

    $configPath = "$env:USERPROFILE\.mambarc"
    $configContent = @"
pkgs_dirs:
   - $env:MAMBA_ROOT_PREFIX\pkgs
"@

    Set-Content -Path $configPath -Value $configContent -Encoding UTF8
    # Write-Host "Configuration file created at $configPath"
}


<#
.SYNOPSIS
    Checks if the specified micromamba environment exists.

.PARAMETER EnvName
    The name of the micromamba environment to check.

.OUTPUTS
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
    $envList = & "$env:MAMBA_ROOT_PREFIX\micromamba.exe" env list | Select-String -Pattern $pattern
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

.OUTPUTS
    [HASHTABLE[]]
    keys:
    - PackageName
    - Success

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
                & "$env:MAMBA_ROOT_PREFIX\micromamba.exe" run -n $EnvName pip install $package --trusted-host pypi.org --trusted-host files.pythonhosted.org *>&1 | Out-Null
            }
            else {
                # Run the install command without trusted hosts
                & "$env:MAMBA_ROOT_PREFIX\micromamba.exe" run -n $EnvName pip install $package *>&1 | Out-Null
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

.OUTPUTS
    [bool] indicates the micromamba environment was created

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
        & "$env:MAMBA_ROOT_PREFIX\micromamba.exe" create -n $EnvName --yes --ssl-verify False python=$PythonVersion pip -c conda-forge *>&1 | Out-Null
    }
    else {
        # redirect output since uipath captures the output
        & "$env:MAMBA_ROOT_PREFIX\micromamba.exe" create -n $EnvName --yes python=$PythonVersion pip -c conda-forge *>&1 | Out-Null
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
    Executes a Python script or a console script within a specified micromamba environment.

.DESCRIPTION
    The `Invoke-PythonScript` function is designed to execute either a Python script (e.g., `script.py`) or 
    a console script (e.g., `detection_coverage`) within a micromamba environment. 
    It uses the `micromamba` command to run the specified script in the specified environment.

.PARAMETER ScriptPath
    The path to the script to execute. 
    - If it ends with `.py`, it is treated as a Python file and executed with the `python` command.
    - Otherwise, it is treated as a console script and executed directly.

.PARAMETER EnvName
    The name of the micromamba environment where the script will be executed.

.PARAMETER Arguments
    An optional array of strings representing the arguments to pass to the script. 
    If not provided, the script will be executed without additional arguments.

.EXAMPLE
    # Execute a Python script within the micromamba environment
    Invoke-PythonScript -ScriptPath "C:\path\to\script.py" -EnvName "langchain" -Arguments @("--option1", "value1")

.EXAMPLE
    # Execute a console script (e.g., detection_coverage) within the micromamba environment
    Invoke-PythonScript -ScriptPath "detection_coverage" -EnvName "langchain" -Arguments @("--help")

.EXAMPLE
    # Execute a script without additional arguments
    Invoke-PythonScript -ScriptPath "detection_coverage" -EnvName "langchain"


.NOTES
    - Ensure that the micromamba environment is properly set up and contains the required scripts or modules.
    - Console scripts must be installed via the `setup.py` `entry_points` configuration.

.OUTPUTS
    Returns the output of the executed script as a string. If an error occurs, the error message is returned.
#>
function Invoke-PythonScript {
    param (
        [Parameter(Mandatory = $true)]
        [string]$ScriptPath,  # Could be a console script or a Python file

        [Parameter(Mandatory = $true)]
        [string]$EnvName,

        [Parameter(Mandatory = $false)]
        [string[]]$Arguments = @() # Optional parameter with a default value
    )

    try {
        # Convert multiline arguments into single-line strings
        $SanitizedArguments = $Arguments | ForEach-Object { $_ -replace "`r?`n", " " }

        # Quote and join the sanitized arguments
        $Command = $SanitizedArguments | ForEach-Object { "`"$($_.ToString())`"" }

        if ($ScriptPath -match '\.py$') {
            # If it's a Python script
            [string]$finalResult = & "$env:MAMBA_ROOT_PREFIX\micromamba.exe" run -n $EnvName python $ScriptPath $Command
        } else {
            # If it's a console script
            [string]$finalResult = & "$env:MAMBA_ROOT_PREFIX\micromamba.exe" run -n $EnvName $ScriptPath $Command
        }

        return [string]$finalResult
    } catch [System.Exception] {
        # Handle errors gracefully
        return $_.Message
    }
}


<#
.SYNOPSIS
    Imports a .env file and loads its key-value pairs as environment variables into the PowerShell runtime.

.DESCRIPTION
    The `Import-DotEnv` function reads a .env file, parses its contents, and sets each key-value pair as an
    environment variable in the current PowerShell session. This allows the environment variables to be accessed
    using `$env:`.

.PARAMETER EnvFilePath
    The path to the .env file. Defaults to the current directory's .env file if not specified.

.EXAMPLE
    # Import a .env file and load its key-value pairs as environment variables
    Import-DotEnv -EnvFilePath "C:\path\to\.env"

.EXAMPLE
    # Import the default .env file in the current directory
    Import-DotEnv

.NOTES
    - The function assumes the .env file uses the standard KEY=VALUE format.
    - If the .env file is not found or cannot be read, the function will return `$false`.

.LINK
    https://12factor.net/config
#>
function Import-DotEnv {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $false)]
        [string]$EnvFilePath = ".env"
    )

    try {
        # Check if the .env file exists
        if (-Not (Test-Path -Path $EnvFilePath)) {
            return $false
        }

        # Read all lines from the .env file
        $envContent = Get-Content -Path $EnvFilePath -ErrorAction Stop

        # Parse each line for key-value pairs
        foreach ($line in $envContent) {
            if ($line -match "^(\w+)=(.+)$") {
                $key = $matches[1]
                $value = $matches[2]

                # Remove surrounding quotes from the value if they exist
                if ($value -match '^"(.*)"$') {
                    $value = $matches[1]
                }

                # Set the environment variable in the current session using Set-Item
                Set-Item -Path "Env:\$key" -Value $value
            }
        }

        return $true
    } catch {
        return $false
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

.OUTPUTS
    [bool] indicating download was successful

.EXAMPLE
    Get-MicromambaBinary
    Downloads, verifies, and extracts the micromamba binary from the default URLs.

.EXAMPLE
    Get-MicromambaBinary -Url "https://example.com/custom/micromamba.exe" -ChecksumUrl "https://example.com/custom/micromamba.sha256"
    Downloads, verifies, and extracts the micromamba binary from custom URLs.
#>
function Get-MicromambaBinary {
    param (
        [string]$Url = "https://github.com/mamba-org/micromamba-releases/releases/latest/download/micromamba-win-64",
        [string]$ChecksumUrl = "https://github.com/mamba-org/micromamba-releases/releases/latest/download/micromamba-win-64.sha256"
    )

    $destinationPath = $env:MAMBA_ROOT_PREFIX
    $binaryPath = Join-Path -Path $destinationPath -ChildPath "micromamba.exe"
    $checksumPath = Join-Path -Path $destinationPath -ChildPath "micromamba-win-64.sha256"

    try {

        if (-not (Test-Path $destinationPath) ) {
            New-Item -Type Directory -Path $destinationPath -Force | Out-Null
        }

        # Download the micromamba binary
        Invoke-WebRequest -Uri $Url -OutFile $binaryPath -UseDefaultCredentials

        # Download the SHA256 checksum file
        Invoke-WebRequest -Uri $ChecksumUrl -OutFile $checksumPath -UseDefaultCredentials

        # Verify both files exist
        if (-not (Test-Path -Path $binaryPath) ) {
            
            throw [System.Exception]::new("Binary path NOT found - $binaryPath")
        }

        if(-not (Test-Path -Path $checksumPath) ) {

            throw [System.Exception]::new("Checksum path NOT found - $checksumPath")
        }

        # Read the checksum from the SHA256 file
        $expectedChecksum = (Get-Content -Path $checksumPath -ErrorAction Stop).Trim()

        # Compute the actual checksum of the downloaded binary
        $actualChecksum = (Get-FileHash -Path $binaryPath -Algorithm SHA256).Hash

        # Compare checksums
        if ($expectedChecksum -ne $actualChecksum) {

            throw [System.Exception]::new("Checksum does NOT match - expected: $expectedChecksum  actual: $actualChecksum")
        }

        # Return success if everything checks out
        return $true
    } catch [Exception] {
        # Write-Host $_.Message

        # Clean up files on error
        Remove-Item -Path $binaryPath -ErrorAction SilentlyContinue
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

.OUTPUTS
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
        & "$env:MAMBA_ROOT_PREFIX\micromamba.exe" env remove -n $EnvName --yes *>&1 | Out-Null
        
        if ($LASTEXITCODE -ne 0) {
            return $false
        }

        # Clean up cached packages
        & "$env:MAMBA_ROOT_PREFIX\micromamba.exe" clean --all --yes *>&1 | Out-Null
        
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

.OUTPUTS
    [bool] indicating if the removal was successful.

.EXAMPLE
    Remove-Micromamba
#>
function Remove-Micromamba {
    try {
        # Define path for micromamba executable
        $micromambaPath = Join-Path -Path $env:MAMBA_ROOT_PREFIX -ChildPath "micromamba.exe"

        # Define path for micromamba root directory
        $micromambaRoot = $env:MAMBA_ROOT_PREFIX

        # Define path for micromamba config file
        $micromambaConfig = Join-Path -Path $env:USERPROFILE -ChildPath ".mambarc"

        # Define path for conda
        $conda = Join-Path -Path $env:USERPROFILE -ChildPath ".conda"

        # Define path for .mamba
        $mamba = Join-Path -Path $env:USERPROFILE -ChildPath ".mamba"


        # Remove micromamba executable
        if (Test-Path -Path $micromambaPath) {
            Remove-Item -Path $micromambaPath -Force -Recurse
        }

        # Remove micromamba root directory
        if (Test-Path -Path $micromambaRoot) {
            Remove-Item -Path $micromambaRoot -Force -Recurse
        }

        # Remove micromamba config file
        if (Test-Path -Path $micromambaConfig) {
            Remove-Item -Path $micromambaConfig -Force
        }

        # Remove conda files
        if (Test-Path -Path $conda) {
            Remove-Item -Path $conda -Force -Recurse
        }

        # Remove mamba files
        if (Test-Path -Path $mamba) {
            Remove-Item -Path $mamba -Force -Recurse
        }

        # Unset the MAMBA_ROOT_PREFIX environment variable
        Remove-Item -Path Env:\MAMBA_ROOT_PREFIX -ErrorAction SilentlyContinue

        return $true
    } catch {
        return $false
    }
}


Export-ModuleMember -Function Import-DotEnv, Remove-Micromamba, Remove-MicromambaEnvironment, Get-MicromambaBinary, Invoke-PythonScript, New-MicromambaEnvironment, Install-PackagesInMicromambaEnvironment, Test-MicromambaEnvironment, Initialize-MambaRootPrefix