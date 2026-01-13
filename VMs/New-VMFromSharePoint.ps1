<#
.SYNOPSIS
    Automates the creation of a Gen 2 Hyper-V virtual machine by downloading a
    multi-part VHDX archive from SharePoint Online, extracting it, and provisioning
    the VM. This script uses MSAL.PS for authentication and native REST API calls.

.DESCRIPTION
    This script performs the following actions:
    1.  Connects to SharePoint Online using the MSAL.PS module for modern authentication.
    2.  Presents an interactive picker to select a folder from the specified library.
    3.  Sets the name of the new VM to match the selected folder's name.
    4.  Downloads all parts of a 7-Zip archive (.7z.001, .7z.002, etc.) from the selected folder.
    5.  Uses 7-Zip (7z.exe) to extract the VHDX file from the downloaded archive.
    6.  Creates a new Gen 2 Hyper-V VM using the extracted VHDX.
    7.  Configures the VM's memory, network switch, and CPU core count.
    8.  Cleans up by deleting the downloaded archive files and the extracted VHDX.

.NOTES
    Author: Gemini Enterprise
    Version: 4.0
    Prerequisites:
        - Hyper-V PowerShell Module
        - MSAL.PS Module (Run: Install-Module MSAL.PS -Force)
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
    
        # Check for MSAL.PS module
        if (-not (Get-Module -ListAvailable -Name MSAL.PS)) {
            throw "The MSAL.PS module is required. Please run 'Install-Module MSAL.PS -Force' from an administrative PowerShell session."
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

        # --- 1. Connect to SharePoint Online via MSAL.PS ---
        Write-Host "Step 1: Authenticating to SharePoint Online..." -ForegroundColor Green
        $spoResourceId = [uri]$spSiteUrl
        $scope = "$($spoResourceId.Scheme)://$($spoResourceId.Host)/.default"
    
        $msalToken = Get-MsalToken -ClientId "1950a258-227b-4e31-a9cf-717495945fc2" -RedirectUri "urn:ietf:wg:oauth:2.0:oob" -Scopes $scope -Interactive
        $authToken = $msalToken.AccessToken
        $headers = @{ "Authorization" = "Bearer $authToken" }
        Write-Host "Successfully obtained authentication token." -ForegroundColor Yellow


        # --- 2. Interactively Select a Folder and Set VM Name ---
        Write-Host "Step 2: Fetching folders from library '$spLibrary' for selection..." -ForegroundColor Green
        $siteRelativeLibraryUrl = ($spSiteUrl -replace [uri]::new($spSiteUrl).GetLeftPart('Authority'), "") + "/" + $spLibrary
        $foldersApiUrl = "$spSiteUrl/_api/web/GetFolderByServerRelativeUrl('$siteRelativeLibraryUrl')/Folders?`$select=Name"
    
        $folderResponse = Invoke-RestMethod -Uri $foldersApiUrl -Headers $headers -Method Get -ErrorAction Stop
        if (-not $folderResponse.value) {
            throw "No folders found in the '$spLibrary' library. Please check the library name and ensure it contains folders."
        }

        $selectedFolder = $folderResponse.value | Out-GridView -Title "Select the Folder Containing the VHDX Archive" -PassThru
    
        if (-not $selectedFolder) {
            throw "No folder was selected. Aborting script."
        }

        $spFolderName = $selectedFolder.Name
        Write-Host "User selected folder: '$spFolderName'" -ForegroundColor Yellow

        # DYNAMICALLY SET VM NAME based on the selected folder
        Write-Host "Setting VM Name based on selected folder..." -ForegroundColor Cyan
        # Sanitize folder name to create a valid VM Name (allow letters, numbers, hyphens, and underscores)
        $vmName = $spFolderName -replace '[^a-zA-Z0-9-_]', ''
        if ([string]::IsNullOrWhiteSpace($vmName)) {
            throw "The sanitized folder name resulted in an empty string ('$spFolderName'). Cannot create VM."
        }
        Write-Host "The new VM will be named: '$vmName'" -ForegroundColor Yellow

        # --- 3. Download Files via REST API ---
        Write-Host "Step 3: Accessing folder '$spFolderName' and downloading files to '$tempPath'..." -ForegroundColor Green
        $folderUrl = "$siteRelativeLibraryUrl/$spFolderName"
        $filesApiUrl = "$spSiteUrl/_api/web/GetFolderByServerRelativeUrl('$folderUrl')/Files?`$select=Name,ServerRelativeUrl"
    
        $filesResponse = Invoke-RestMethod -Uri $filesApiUrl -Headers $headers -Method Get -ErrorAction Stop
        if (-not $filesResponse.value) {
            throw "No files found in folder '$spFolderName'."
        }
        Write-Host "Found $($filesResponse.value.Count) files to download." -ForegroundColor Yellow

        foreach ($file in $filesResponse.value) {
            $targetPath = Join-Path -Path $tempPath -ChildPath $file.Name
            $fileDownloadUrl = "$spSiteUrl/_api/web/GetFileByServerRelativeUrl('$($file.ServerRelativeUrl)')/`$value"
        
            Write-Host "Downloading $($file.Name)..."
            Invoke-RestMethod -Uri $fileDownloadUrl -Headers $headers -Method Get -OutFile $targetPath -ErrorAction Stop
        }
        Write-Host "All files downloaded successfully." -ForegroundColor Yellow

        # --- 4. Extract the VHDX ---
        Write-Host "Step 4: Extracting VHDX from the archive..." -ForegroundColor Green
        $firstArchiveFile = Get-ChildItem -Path $tempPath -Filter "*.7z.001" | Select-Object -First 1
        if (-not $firstArchiveFile) {
            throw "The first part of the 7-Zip archive (*.7z.001) was not found in '$tempPath'."
        }

        Write-Host "Starting extraction on '$($firstArchiveFile.FullName)'..."
        $arguments = @("e", "`"$($firstArchiveFile.FullName)`"", "-o`"$($tempPath)`"", "-y")
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
    }
    finally {
        # --- 6. Cleanup ---
        Write-Host "Step 6: Performing cleanup..." -ForegroundColor Green
    
        if (Test-Path $tempPath) {
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