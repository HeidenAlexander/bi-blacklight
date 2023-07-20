if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Start-Process pwsh -Verb RunAs "-NoProfile -ExecutionPolicy Bypass -Command `"cd '$pwd'; & '$PSCommandPath';`"";
    exit;
}

# Function to check if the destination folder exists and create it if it doesn't
function ClearAndCreate-DestinationFolder {
    param (
        [Parameter(Mandatory=$true)]
        [string]$destinationFolder
    )

    # Check if the folder exists
    if (Test-Path -Path $destinationFolder -PathType Container) {
    # Get all the files and subfolders in the specified folder
    $items = Get-ChildItem -Path $destinationFolder -Force

    # Delete each file
    foreach ($item in $items) {
        if ($item.PSIsContainer) {
            # Delete subfolders and their contents recursively
            Remove-Item -Path $item.FullName -Recurse -Force
        } else {
            # Delete files
            Remove-Item -Path $item.FullName -Force
        }
    }
    } else {
            # Create the destination folder if it doesn't exist
            New-Item -ItemType Directory -Path $destinationFolder | Out-Null
            Write-Host "Destination folder created: $destinationFolder"
    }
}

# Function to check if the destination folder exists and create it if it doesn't
function Create-DestinationFolder {
    param (
        [Parameter(Mandatory=$true)]
        [string]$destinationFolder
    )

    # Check if the destination folder exists
    if (-not (Test-Path -Path $destinationFolder)) {
        # Create the destination folder if it doesn't exist
        New-Item -ItemType Directory -Path $destinationFolder | Out-Null
        Write-Host "Destination folder created: $destinationFolder"
    }
}

# Prompt the user to continue or exit
$choice = Read-Host "Preparing to install BI-Blacklight into C:\Program Files\BI-Blacklight. Continue? (Y/N)"

# Convert the choice to uppercase for case-insensitive comparison
$choice = $choice.ToUpper()

# Check the user's choice
if ($choice -eq "Y") {
    Write-Host "Installing..."

    # Get the current directory as the source folder
    $rootFolder = $PSScriptRoot #Split-Path -Path $script:MyInvocation.MyCommand.Path

    # Specify the source items (files or folders) to be copied
    $sourceItems = @(
        "$rootFolder\Application\Run.bat",
        "$rootFolder\Report Setup",
        "$rootFolder\Application\Workflow",
        "$rootFolder\Application\Install_Packages.bat",
        "$rootFolder\LICENSE",
        "$rootFolder\README.md"
    )

    # Specify external tool registration file to copy
    $registrationFile = "$rootFolder\Application\BI-Blacklight.pbitool.json"

    # Specify the destination folder where the files will be copied
    $destinationFolder = "C:\Program Files\BI-Blacklight"
    $externalToolsFolder = "C:\Program Files (x86)\Common Files\Microsoft Shared\Power BI Desktop\External Tools"


    # Check if the destination install folder exists and delete contents or create it if necessary
    ClearAndCreate-DestinationFolder -destinationFolder $destinationFolder


    # Check if the external tools folder exists and create it if necessary
    Create-DestinationFolder -destinationFolder $externalToolsFolder

    # Copy items from the source to the install folder
    $sourceItems | ForEach-Object {
        $sourceItem = $_
        $destinationPath = Join-Path -Path $destinationFolder -ChildPath (Split-Path -Path $sourceItem -Leaf)
        Copy-Item -Path $sourceItem -Destination $destinationPath -Force -Recurse
    }

    # Copy external tool registration file to external tools folder
    Copy-Item -Path $registrationFile -Destination $externalToolsFolder -Force -Recurse

    # Check if the copy operation was successful
    if ($?) {
        Write-Host "Files copied successfully."
    }
    else {
        Write-Host "An error occurred while copying files."
    }
}
elseif ($choice -eq "N") {
    Write-Host "Exiting the script..."
    exit
}
else {
    Write-Host "Invalid choice. Exiting the script..."
    exit
}
# Prompt the user to install Python packages or exit.
$choice = Read-Host "Do you want to install required Python packages? (Y/N)"
$choice = $choice.ToUpper()

# Check the user's choice
if ($choice -eq "Y") {
    Write-Host "Continuing with the script..."

    Write-Host "Installing required Python packages."

    $batchFilePath = "C:\Program Files\BI-Blacklight\Install_Packages.bat"
    $process = Start-Process -FilePath "cmd.exe" -ArgumentList "/c `"$batchFilePath`"" -PassThru

    # Wait for the batch file to finish executing
    $process.WaitForExit()

    # Check the exit code of the batch file
    $exitCode = $process.ExitCode

    # Display the exit code
    Write-Host "Batch file exit code: $exitCode"
}
elseif ($choice -eq "N") {
    Write-Host "Exiting the script..."
    exit
}
else {
    Write-Host "Invalid choice. Exiting the script..."
    exit
}

Read-Host -Prompt "Completed. Press Enter to exit"