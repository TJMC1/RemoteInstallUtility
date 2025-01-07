Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# At the start of your script, define the PsExec path relative to the script location
$psExecPath = Join-Path $PSScriptRoot "tools\PsExec.exe"

# Verify PsExec exists locally
if (-not (Test-Path $psExecPath)) {
    throw "PsExec not found at: $psExecPath. Please ensure PsExec.exe is in the 'tools' folder."
}

# Create the main form
$form = New-Object System.Windows.Forms.Form
$form.Text = 'Remote Installation Manager'
$form.Size = New-Object System.Drawing.Size(800,600)
$form.StartPosition = 'CenterScreen'

# Server List File Selection
$lblServerFile = New-Object System.Windows.Forms.Label
$lblServerFile.Location = New-Object System.Drawing.Point(10,20)
$lblServerFile.Size = New-Object System.Drawing.Size(100,20)
$lblServerFile.Text = 'Server List File:'
$form.Controls.Add($lblServerFile)

$txtServerFile = New-Object System.Windows.Forms.TextBox
$txtServerFile.Location = New-Object System.Drawing.Point(120,20)
$txtServerFile.Size = New-Object System.Drawing.Size(550,20)
$form.Controls.Add($txtServerFile)

$btnBrowseServers = New-Object System.Windows.Forms.Button
$btnBrowseServers.Location = New-Object System.Drawing.Point(680,20)
$btnBrowseServers.Size = New-Object System.Drawing.Size(100,20)
$btnBrowseServers.Text = 'Browse'
$form.Controls.Add($btnBrowseServers)

# Installer Path Selection
$lblInstallerPath = New-Object System.Windows.Forms.Label
$lblInstallerPath.Location = New-Object System.Drawing.Point(10,50)
$lblInstallerPath.Size = New-Object System.Drawing.Size(100,20)
$lblInstallerPath.Text = 'Installer Path:'
$form.Controls.Add($lblInstallerPath)

$txtInstallerPath = New-Object System.Windows.Forms.TextBox
$txtInstallerPath.Location = New-Object System.Drawing.Point(120,50)
$txtInstallerPath.Size = New-Object System.Drawing.Size(550,20)
$form.Controls.Add($txtInstallerPath)

$btnBrowseInstaller = New-Object System.Windows.Forms.Button
$btnBrowseInstaller.Location = New-Object System.Drawing.Point(680,50)
$btnBrowseInstaller.Size = New-Object System.Drawing.Size(100,20)
$btnBrowseInstaller.Text = 'Browse'
$form.Controls.Add($btnBrowseInstaller)

# Installer Arguments
$lblInstallerArgs = New-Object System.Windows.Forms.Label
$lblInstallerArgs.Location = New-Object System.Drawing.Point(10,80)
$lblInstallerArgs.Size = New-Object System.Drawing.Size(100,20)
$lblInstallerArgs.Text = 'Installer Args:'
$form.Controls.Add($lblInstallerArgs)

$txtInstallerArgs = New-Object System.Windows.Forms.TextBox
$txtInstallerArgs.Location = New-Object System.Drawing.Point(120,80)
$txtInstallerArgs.Size = New-Object System.Drawing.Size(550,20)
$txtInstallerArgs.Text = "/silent"  # Default value
$form.Controls.Add($txtInstallerArgs)

# Status ListView
$listView = New-Object System.Windows.Forms.ListView
$listView.Location = New-Object System.Drawing.Point(10,120)
$listView.Size = New-Object System.Drawing.Size(770,400)
$listView.View = [System.Windows.Forms.View]::Details
$listView.FullRowSelect = $true
$listView.GridLines = $true

# Add columns
$listView.Columns.Add("Status", 50)  # Small column for status circle
$listView.Columns.Add("Server", 300)  # Made wider since we removed time columns
$listView.Columns.Add("Message", 300)
$listView.Columns.Add("Progress", 100)
$form.Controls.Add($listView)

# Create status images
$statusImageList = New-Object System.Windows.Forms.ImageList
$statusImageList.ImageSize = New-Object System.Drawing.Size(16, 16)

# Create status circles
foreach ($color in @("Red", "Yellow", "Green")) {
    $bitmap = New-Object System.Drawing.Bitmap(16, 16)
    $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
    $brush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::$color)
    $graphics.FillEllipse($brush, 0, 0, 15, 15)
    $statusImageList.Images.Add($color.ToLower(), $bitmap)
    $brush.Dispose()
    $graphics.Dispose()
}

$listView.SmallImageList = $statusImageList

# Start Installation Button
$btnStart = New-Object System.Windows.Forms.Button
$btnStart.Location = New-Object System.Drawing.Point(10,530)
$btnStart.Size = New-Object System.Drawing.Size(770,30)
$btnStart.Text = 'Start Installation'
$form.Controls.Add($btnStart)

# Add browse button click handlers
$btnBrowseServers.Add_Click({
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*"
    if ($openFileDialog.ShowDialog() -eq 'OK') {
        $txtServerFile.Text = $openFileDialog.FileName
    }
})

$btnBrowseInstaller.Add_Click({
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Filter = "Executable files (*.exe)|*.exe|All files (*.*)|*.*"
    if ($openFileDialog.ShowDialog() -eq 'OK') {
        $txtInstallerPath.Text = $openFileDialog.FileName
    }
})

# Function to enable/disable all GUI controls during installation
function Set-ControlState {
    param([bool]$enabled)
    $btnStart.Enabled = $enabled
    $btnBrowseServers.Enabled = $enabled
    $btnBrowseInstaller.Enabled = $enabled
    $txtServerFile.Enabled = $enabled
    $txtInstallerPath.Enabled = $enabled
    $txtInstallerArgs.Enabled = $enabled
}

# Main button click event handler for starting installations
$btnStart.Add_Click({
    # Validate that both server list and installer path are provided
    if (-not $txtServerFile.Text -or -not $txtInstallerPath.Text) {
        [System.Windows.Forms.MessageBox]::Show("Please select both server list and installer files.", "Error")
        return
    }

    # Disable GUI controls while installation is running
    Set-ControlState -enabled $false

    # Clear previous results and initialize variables
    $listView.Items.Clear()
    $servers = Get-Content $txtServerFile.Text
    $installer = $txtInstallerPath.Text
    $script:runningJobs = @{}  # Hashtable to track installation jobs

    # Create a ListView item for each server
    foreach ($server in $servers) {
        $item = New-Object System.Windows.Forms.ListViewItem
        $item.ImageKey = "yellow"  # Start with yellow (in progress)
        $item.SubItems.Add($server)
        $item.SubItems.Add("Starting")
        $item.SubItems.Add("0%")
        $listView.Items.Add($item)

        # Start a background job for each server installation
        $job = Start-Job -ScriptBlock {
            param($server, $installer, $installerArgs, $psExecPath)
            try {
                # Create remote directory if it doesn't exist
                $remotePath = "C:\Temp\remoteinstallutility"
                Invoke-Command -ComputerName $server -ScriptBlock {
                    param($path)
                    if (-not (Test-Path $path)) {
                        New-Item -ItemType Directory -Path $path -Force | Out-Null
                    }
                } -ArgumentList $remotePath

                # Copy installer to remote server using admin share
                $installerName = Split-Path $installer -Leaf
                $remoteInstaller = Join-Path $remotePath $installerName
                Copy-Item -Path $installer -Destination "\\$server\c$\Temp\remoteinstallutility\$installerName" -Force

                # Create remote directory and copy PsExec before main execution
                $remotePsExecDir = "\\$server\c$\Temp\RemoteInstallUtility\tools"
                if (-not (Test-Path $remotePsExecDir)) {
                    New-Item -ItemType Directory -Path $remotePsExecDir -Force | Out-Null
                }
                Copy-Item -Path $psExecPath -Destination $remotePsExecDir -Force

                # Execute the installer on the remote server with custom arguments
                $result = Invoke-Command -ComputerName $server -ScriptBlock {
                    param($installerPath, $installerArgs, $psExecPath)
                    try {
                        # Define root directory and subfolders
                        $rootDir = "C:\Temp\RemoteInstallUtility"
                        $directories = @{
                            Root = $rootDir
                            Tools = Join-Path $rootDir "tools"
                            Logs = Join-Path $rootDir "logs"
                        }

                        # Create directory structure
                        foreach ($dir in $directories.Values) {
                            if (-not (Test-Path $dir)) {
                                New-Item -ItemType Directory -Path $dir -Force | Out-Null
                            }
                        }
                        
                        # Copy PsExec if it's not already there
                        $remotePsExecPath = Join-Path $directories.Tools "PsExec.exe"
                        if (-not (Test-Path $remotePsExecPath)) {
                            Copy-Item -Path $psExecPath -Destination $remotePsExecPath -Force
                        }

                        # Create timestamp for logging
                        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
                        $installerName = [System.IO.Path]::GetFileNameWithoutExtension($installerPath)
                        $logPath = Join-Path $directories.Logs "$($installerName)_$timestamp.log"

                        Write-Output "Starting installation process..." | Out-File $logPath
                        Write-Output "Installer path: $installerPath" | Out-File $logPath -Append
                        Write-Output "PsExec path: $remotePsExecPath" | Out-File $logPath -Append
                        Write-Output "Arguments: $installerArgs" | Out-File $logPath -Append
                        
                        # Verify files exist
                        Write-Output "Verifying files..." | Out-File $logPath -Append
                        Write-Output "Installer exists: $(Test-Path $installerPath)" | Out-File $logPath -Append
                        Write-Output "PsExec exists: $(Test-Path $remotePsExecPath)" | Out-File $logPath -Append

                        # Create full PsExec command
                        $psExecArgs = "-s -i -d -accepteula -h cmd /c `"cd /d $($directories.Root) && `"$installerPath`" $installerArgs`""
                        Write-Output "Full PsExec command: $remotePsExecPath $psExecArgs" | Out-File $logPath -Append
                        
                        # Launch installer with PsExec through cmd
                        $process = Start-Process -FilePath $remotePsExecPath `
                            -ArgumentList $psExecArgs `
                            -PassThru `
                            -NoNewWindow `
                            -WorkingDirectory $directories.Root `
                            -RedirectStandardOutput (Join-Path $directories.Logs "psexec_output_$timestamp.log") `
                            -RedirectStandardError (Join-Path $directories.Logs "psexec_error_$timestamp.log")
                        
                        if ($process -eq $null) {
                            throw "Failed to start process"
                        }

                        Write-Output "Process started with ID: $($process.Id)" | Out-File $logPath -Append
                        Write-Output "Waiting for process to start..." | Out-File $logPath -Append
                        Start-Sleep -Seconds 2

                        # Check if process is still running
                        $runningProcess = Get-Process -Id $process.Id -ErrorAction SilentlyContinue
                        Write-Output "Process still running: $($runningProcess -ne $null)" | Out-File $logPath -Append

                        return @{
                            Status = "Running"
                            ProcessId = $process.Id
                            Path = $installerPath
                            LogPath = $logPath
                            RootDir = $rootDir
                        }
                    }
                    catch {
                        Write-Output "Error: $_" | Out-File $logPath -Append
                        Write-Output "Stack trace: $($_.ScriptStackTrace)" | Out-File $logPath -Append
                        throw "Failed to start installer: $_"
                    }
                } -ArgumentList $remoteInstaller, $installerArgs, $psExecPath

                return $result
            }
            catch {
                return @{
                    Status = "Failed"
                    Error = $_.Exception.Message
                }
            }
        } -ArgumentList $server, $installer, $txtInstallerArgs.Text, $psExecPath

        # Store the job reference for status checking
        $script:runningJobs[$server] = $job
    }

    # Create a timer to periodically check installation status
    $script:timer = New-Object System.Windows.Forms.Timer
    $script:timer.Interval = 5000 # Check every 5 seconds
    $script:completionMessageShown = $false

    # Modify the timer tick section
    $script:timer.Add_Tick({
        $allCompleted = $true
        $currentJobs = @() + $script:runningJobs.Keys  # Create a copy of keys to enumerate

        foreach ($server in $currentJobs) {
            $job = $script:runningJobs[$server]
            $item = $listView.Items | Where-Object { $_.SubItems[1].Text -eq $server }

            if ($job.State -eq 'Completed') {
                $result = Receive-Job -Job $job
                
                try {
                    # Check if the installation process is still running on remote server
                    if ($result.ProcessId) {
                        $processStatus = Invoke-Command -ComputerName $server -ScriptBlock {
                            param($processId, $installerPath, $logPath)
                            try {
                                $process = Get-Process -Id $processId -ErrorAction SilentlyContinue
                                
                                # Read the log file if it exists
                                $logContent = if (Test-Path $logPath) { 
                                    Get-Content $logPath -Raw 
                                } else { 
                                    "No log file found at $logPath" 
                                }
                                
                                if ($process) { 
                                    return @{
                                        Status = "Running"
                                        Details = "Process still running with ID: $processId"
                                        Log = $logContent
                                    }
                                }
                                
                                # If process is no longer running, assume it completed
                                # You might want to add additional checks here specific to your needs
                                return @{
                                    Status = "Completed"
                                    Details = "Process completed"
                                    Log = $logContent
                                }
                            }
                            catch {
                                return @{
                                    Status = "Error"
                                    Details = "Error checking process: $_"
                                    Log = if (Test-Path $logPath) { Get-Content $logPath -Raw } else { "No log file found" }
                                }
                            }
                        } -ArgumentList $result.ProcessId, $result.Path, $result.LogPath

                        # Update status based on process state
                        switch ($processStatus.Status) {
                            "Running" {
                                $item.ImageKey = "yellow"
                                $item.SubItems[2].Text = "Installing: $($processStatus.Details)"
                                $allCompleted = $false
                            }
                            "Completed" {
                                $item.ImageKey = "green"
                                $item.SubItems[2].Text = "Completed. Check log at: $($result.LogPath)"
                                $item.SubItems[3].Text = "100%"
                                Remove-Job -Job $job
                                $script:runningJobs.Remove($server)
                            }
                            default {
                                $item.ImageKey = "red"
                                $item.SubItems[2].Text = "Failed: $($processStatus.Details). Check log at: $($result.LogPath)"
                                $item.SubItems[3].Text = "0%"
                                Remove-Job -Job $job
                                $script:runningJobs.Remove($server)
                            }
                        }
                    } else {
                        $item.ImageKey = "red"
                        $item.SubItems[2].Text = "Failed: No process ID returned"
                        $item.SubItems[3].Text = "0%"
                        Remove-Job -Job $job
                        $script:runningJobs.Remove($server)
                    }
                }
                catch {
                    $item.ImageKey = "red"
                    $item.SubItems[2].Text = "Failed: $($result.Error)"
                    $item.SubItems[3].Text = "0%"
                    Remove-Job -Job $job
                    $script:runningJobs.Remove($server)
                }
            }
            else {
                $allCompleted = $false
            }
        }

        if ($allCompleted -and -not $script:completionMessageShown) {
            $script:timer.Stop()
            Set-ControlState -enabled $true
            $script:completionMessageShown = $true
            [System.Windows.Forms.MessageBox]::Show("All installations completed!", "Complete")
        }
    })
    $script:timer.Start()
})

# Show the form
$form.ShowDialog()