if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Start-Process pwsh -Verb RunAs "-NoProfile -ExecutionPolicy Bypass -Command `"cd '$pwd'; & '$PSCommandPath';`"";
    exit;
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

# Get the current directory as the source folder
$rootFolder = $PSScriptRoot #Split-Path -Path $script:MyInvocation.MyCommand.Path

# Specify the source items (files or folders) to be copied
$sourceItems = @(
    "$rootFolder\Run.bat",
    "$rootFolder\Report Setup",
    "$rootFolder\Workflow",
    "$rootFolder\Install_Packages.bat",
    "$rootFolder\LICENSE",
    "$rootFolder\README.md"
)

# Specify external tool registration file to copy
$registrationFile = "$rootFolder\BI-Blacklight.pbitool.json"

# Specify the destination folder where the files will be copied
$destinationFolder = "C:\Program Files\BI-Blacklight"
$externalToolsFolder = "C:\Program Files (x86)\Common Files\Microsoft Shared\Power BI Desktop\External Tools"


# Check if the destination install folder exists and create it if necessary
Create-DestinationFolder -destinationFolder $destinationFolder

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
} else {
    Write-Host "An error occurred while copying files."
}

Read-Host -Prompt "Press Enter to exit"