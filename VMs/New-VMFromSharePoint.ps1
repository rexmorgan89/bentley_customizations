<#
.SYNOPSIS
    Automates the creation of a Gen 2 Hyper-V virtual machine by downloading a
    multi-part VHDX archive from SharePoint Online, extracting it, and provisioning
    the VM.

.DESCRIPTION
    This script performs the following actions:
    1.  Connects to a specified SharePoint Online site using PnP.PowerShell.
    2.  Presents an interactive picker to select a folder from the specified library.
    3.  Downloads all parts of a 7-Zip archive (.7z.001, .7z.002, etc.) from the selected folder.
    4.  Uses 7-Zip (7z.exe) to extract the VHDX file from the downloaded archive.
    5.  Creates a new Gen 2 Hyper-V VM using the extracted VHDX.
    6.  Configures the VM's memory, network switch, and CPU core count.
    7.  Cleans up by deleting the downloaded archive files and the extracted VHDX.

.NOTES
    Author: Gemini Enterprise
    Version: 2.0
    Prerequisites:
        - Hyper-V PowerShell Module
        - PnP.PowerShell Module (Install-Module PnP.PowerShell)
        - 7-Zip installed and 7z.exe accessible via the system's PATH or a direct path.
#>

[CmdletBinding()]
param(
   
)

process {
    
    #==============================================================================
    # SCRIPT LOGIC - DO NOT MODIFY BELOW THIS LINE
    #==============================================================================

    # --- Main Execution Block ---
    try {
        # Check for admin privileges, required for Hyper-V cmdlets
        if (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
            throw "This script must be run with Administrator privileges."
        }

        # Ensure the 7-Zip executable exists
        if (-NOT (Test-Path $sevenZipPath)) {
            throw "7-Zip executable not found at '$sevenZipPath'. Please update the \$sevenZipPath variable."
        }

        # Create the temporary directory if it doesn't exist
        if (-NOT (Test-Path $tempPath)) {
            Write-Host "Creating temporary directory at '$tempPath'..." -ForegroundColor Cyan
            New-Item -Path $tempPath -ItemType Directory | Out-Null
        }

        # --- 1. Connect to SharePoint Online ---
        Write-Host "Step 1: Connecting to SharePoint site '$spSiteUrl'..." -ForegroundColor Green
        Connect-PnPOnline -Url $spSiteUrl -Interactive -ErrorAction Stop
        $web = Get-PnPWeb
        Write-Host "Successfully connected to site: $($web.Title)" -ForegroundColor Yellow

        # --- 2. Select Folder and List Files ---
        Write-Host "Step 2: Accessing folder '$spFolderName' in library '$spLibrary'..." -ForegroundColor Green
        $folderUrl = "$spLibrary/$spFolderName"
        $files = Get-PnPFolderItem -FolderSiteRelativeUrl $folderUrl -ItemType File -ErrorAction Stop
        if (-not $files) {
            throw "No files found in folder '$folderUrl'. Please check the path."
        }
        Write-Host "Found $($files.Count) files to download." -ForegroundColor Yellow

        # --- 3. Download Files ---
        Write-Host "Step 3: Downloading files to '$tempPath'..." -ForegroundColor Green
        foreach ($file in $files) {
            $targetPath = Join-Path -Path $tempPath -ChildPath $file.Name
            Write-Host "Downloading $($file.Name)..."
            Get-PnPFile -ServerRelativeUrl $file.ServerRelativeUrl -Path $tempPath -Filename $file.Name -AsFile -Force -ErrorAction Stop
        }
        Write-Host "All files downloaded successfully." -ForegroundColor Yellow

        # --- 4. Extract the VHDX ---
        Write-Host "Step 4: Extracting VHDX from the archive..." -ForegroundColor Green
        $firstArchiveFile = Get-ChildItem -Path $tempPath -Filter "*.7z.001" | Select-Object -First 1
        if (-not $firstArchiveFile) {
            throw "The first part of the 7-Zip archive (*.7z.001) was not found in '$tempPath'."
        }

        Write-Host "Starting extraction on '$($firstArchiveFile.FullName)'..."
        $arguments = @(
            "e", # Extract command
            "`"$($firstArchiveFile.FullName)`"",
            "-o`"$($tempPath)`"", # Output directory
            "-y" # Assume Yes to all queries
        )

        $process = Start-Process -FilePath $sevenZipPath -ArgumentList $arguments -Wait -PassThru -NoNewWindow
        if ($process.ExitCode -ne 0) {
            throw "7-Zip extraction failed with exit code $($process.ExitCode)."
        }
    
        $extractedVhdx = Get-ChildItem -Path $tempPath -Filter "*.vhdx" | Select-Object -First 1
        if (-not $extractedVhdx) {
            throw "Extraction completed, but no VHDX file was found in '$tempPath'."
        }
        Write-Host "Successfully extracted '$($extractedVhdx.Name)'." -ForegroundColor Yellow


        # --- 5. Create the Hyper-V VM ---
        Write-Host "Step 5: Creating Hyper-V virtual machine '$vmName'..." -ForegroundColor Green
    
        # Check if a VM with the same name already exists
        if (Get-VM -Name $vmName -ErrorAction SilentlyContinue) {
            throw "A virtual machine named '$vmName' already exists. Please choose a different name or remove the existing VM."
        }

        Write-Host "Provisioning new VM..."
        New-VM -Name $vmName -MemoryStartupBytes $vmMemory -Generation 2 -VHDPath $extractedVhdx.FullName -SwitchName $vmSwitch -ErrorAction Stop
    
        Write-Host "Configuring VM settings..."
        Set-VM -Name $vmName -ProcessorCount $vmCpuCount -ErrorAction Stop
    
        Write-Host "VM '$vmName' created successfully with the following specifications:" -ForegroundColor Yellow
        Get-VM -Name $vmName | Select-Object VMName, State, Generation, MemoryStartup, ProcessorCount | Format-List


    }
    catch {
        Write-Error "An error occurred: $($_.Exception.Message)"
        # You might want to add more specific cleanup here if the script fails mid-way
    }
    finally {
        # --- 6. Cleanup ---
        Write-Host "Step 6: Performing cleanup..." -ForegroundColor Green
    
        # Disconnect from SharePoint session
        if (Get-PnPConnection -ErrorAction SilentlyContinue) {
            Write-Host "Disconnecting from SharePoint..."
            Disconnect-PnPOnline
        }

        if (Test-Path $tempPath) {
            # Get reference to extracted VHDX again, in case it was created successfully
            $vhdxFileForCleanup = Get-ChildItem -Path $tempPath -Filter "*.vhdx" | Select-Object -First 1

            if ($vhdxFileForCleanup -and (Get-VM -Name $vmName -ErrorAction SilentlyContinue)) {
                Write-Host "VM created. Deleting downloaded 7-Zip files and the original VHDX file..."
                Remove-Item -Path (Join-Path -Path $tempPath -ChildPath "*.7z*") -Force -ErrorAction SilentlyContinue
                Remove-Item -Path $vhdxFileForCleanup.FullName -Force -ErrorAction SilentlyContinue
            }
            else {
                Write-Warning "VM was not created or found. Deleting the entire temp directory for a full cleanup."
                Remove-Item -Path $tempPath -Recurse -Force -ErrorAction SilentlyContinue
            }
            Write-Host "Cleanup complete." -ForegroundColor Yellow
        }
    }

    Write-Host "Script execution finished." -ForegroundColor Green

}

begin {
    #==============================================================================
    # SCRIPT CONFIGURATION - UPDATE THESE VARIABLES
    #==============================================================================

    # --- SharePoint and File Configuration ---
    $spSiteUrl = "https://bentley.sharepoint.com/sites/ProjectWiseHyperVLibrary"
    $spLibrary = "Documents/General" # The name of the Document Library containing the image folders

    # --- Local Machine Configuration ---
    $tempPath = "C:\Temp\VMCreation" # A temporary directory for downloads and extraction
    $sevenZipPath = "C:\Program Files\7-Zip\7z.exe" # Path to the 7-Zip executable

    # --- Hyper-V VM Configuration ---
    $vmName = "MyNewVM"
    $vmMemory = 8GB
    $vmSwitch = "Default Switch"
    $vmCpuCount = 10

    $InformationPreference = 'Continue'
    $VerbosePreference = 'Continue' # Uncomment this line if you want to see verbose messages.

    # Log all script output to a file for easy reference later if needed.
    [string] $lastRunLogFilePath = "$PSCommandPath.LastRun.log"
    Start-Transcript -Path $lastRunLogFilePath

    # Display the time that this script started running.
    [DateTime] $startTime = Get-Date
    Write-Information "Starting script at '$($startTime.ToString('u'))'."
}

end {
    # Display the time that this script finished running, and how long it took to run.
    [DateTime] $finishTime = Get-Date
    [TimeSpan] $elapsedTime = $finishTime - $startTime
    Write-Information "Finished script at '$($finishTime.ToString('u'))'. Took '$elapsedTime' to run."

    Stop-Transcript
}