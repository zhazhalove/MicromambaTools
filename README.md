# Micromamba PowerShell Module

This PowerShell module simplifies the management of Python environments using Micromamba. Follow the steps below to set up and use the module effectively.

## Installation

1. Install the module into the Windows user module folder:
   ```
   C:\Users\<user>\Documents\WindowsPowerShell\Modules\
   ```

## Example Usage

### Step 1: Get Micromamba Binary

Download the Micromamba binary using the following command:

```powershell
Get-MicromambaBinary
```

### Step 2: Initialize the Mamba Root Prefix

Initialize the Micromamba root prefix directory:

```powershell
Initialize-MambaRootPrefix
```

### Step 3: Create a New Micromamba Environment

Create a new environment named `langchain` with Python version 3.11:

```powershell
New-MicromambaEnvironment -EnvName "langchain" -PythonVersion "3.11"
```

### Step 4: Test the Micromamba Environment

Verify the created environment:

```powershell
Test-MicromambaEnvironment -EnvName "langchain"
```

### Step 5: Install Packages in the Micromamba Environment

Install required Python packages in the `langchain` environment:

```powershell
Install-PackagesInMicromambaEnvironment -EnvName "langchain" -Packages @('langchain', 'langchain-openai', 'typer', 'python-dotenv')
```

### Step 6: Run a Python Script

Invoke a Python script within the `langchain` environment with specific arguments:

```powershell
Invoke-PythonScript -ScriptPath "C:\Users\some\path\langchain_organizational_alignment_openai.py" -EnvName "langchain" -Arguments "-i IOCs Multiple alerts generated for an attack technique that has been extensively documented and for which comprehensive mitigation measures are already in place. Threat Intelligence Detailed reports and analysis from cybersecurity communities that have thoroughly covered the attack technique, providing clear guidelines for detection, prevention, and response. Scenario The SIEM system generates numerous alerts related to an attempted SQL injection attack against a web application. This technique is well-documented, and the organization has implemented strong input validation and parameterized queries as recommended mitigations. Despite the high volume of alerts, further investigation confirms that the attempted attacks were effectively neutralized by existing defenses. This scenario underscores the importance of fine-tuning SIEM rules and alert thresholds to focus on emerging threats and reduce the noise from well-understood and adequately mitigated techniques, thereby allowing the security team to allocate resources more efficiently."
```

## License

This project is licensed under the [MIT License](LICENSE).