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
        $item = New-Object System.Windows.Forms.ListViewItem($server)
        $item.SubItems.Add("Starting")
        $item.SubItems.Add("0%")
        $item.SubItems.Add((Get-Date).ToString())
        $item.SubItems.Add("")
        $listView.Items.Add($item)

        # Start a background job for each server installation
        $job = Start-Job -ScriptBlock {
            param($server, $installer, $installerArgs)
            try {
                # Create remote directory if it doesn't exist
                $remotePath = "C:\temp\remoteinstallutility"
                Invoke-Command -ComputerName $server -ScriptBlock {
                    param($path)
                    if (-not (Test-Path $path)) {
                        New-Item -ItemType Directory -Path $path -Force | Out-Null
                    }
                } -ArgumentList $remotePath

                # Copy installer to remote server using admin share
                $installerName = Split-Path $installer -Leaf
                $remoteInstaller = Join-Path $remotePath $installerName
                Copy-Item -Path $installer -Destination "\\$server\c$\temp\remoteinstallutility\$installerName" -Force

                # Execute the installer on the remote server with custom arguments
                $result = Invoke-Command -ComputerName $server -ScriptBlock {
                    param($installerPath, $installerArgs)
                    try {
                        # Start the installer with custom arguments
                        $process = Start-Process -FilePath $installerPath -ArgumentList $installerArgs -PassThru
                        return @{
                            Status = "Running"
                            ProcessId = $process.Id
                        }
                    }
                    catch {
                        throw $_
                    }
                } -ArgumentList $remoteInstaller, $installerArgs

                return $result
            }
            catch {
                return @{
                    Status = "Failed"
                    Error = $_.Exception.Message
                }
            }
        } -ArgumentList $server, $installer, $txtInstallerArgs.Text

        # Store the job reference for status checking
        $script:runningJobs[$server] = $job
    }

    # Create a timer to periodically check installation status
    $timer = New-Object System.Windows.Forms.Timer
    $timer.Interval = 5000 # Check every 5 seconds
    $timer.Add_Tick({
        $allCompleted = $true

        # Check status of each server's installation
        foreach ($server in $script:runningJobs.Keys) {
            $job = $script:runningJobs[$server]
            $item = $listView.Items | Where-Object { $_.Text -eq $server }

            if ($job.State -eq 'Completed') {
                $result = Receive-Job -Job $job
                
                try {
                    # Check if the installation process is still running on remote server
                    $processStatus = Invoke-Command -ComputerName $server -ScriptBlock {
                        param($processId)
                        $process = Get-Process -Id $processId -ErrorAction SilentlyContinue
                        if ($process) { return "Running" }
                        return "Completed"
                    } -ArgumentList $result.ProcessId

                    # Update status based on process state
                    if ($processStatus -eq "Running") {
                        $item.SubItems[1].Text = "Installing"
                        $allCompleted = $false
                    } else {
                        $item.SubItems[1].Text = "Completed"
                        $item.SubItems[2].Text = "100%"
                        $item.SubItems[4].Text = (Get-Date).ToString()
                        Remove-Job -Job $job
                        $script:runningJobs.Remove($server)
                    }
                }
                catch {
                    # Handle failed installations
                    $item.SubItems[1].Text = "Failed: $($result.Error)"
                    $item.SubItems[2].Text = "0%"
                    Remove-Job -Job $job
                    $script:runningJobs.Remove($server)
                }
            }
            else {
                $allCompleted = $false
            }
        }

        # If all installations are complete, clean up and notify user
        if ($allCompleted) {
            $timer.Stop()
            Set-ControlState -enabled $true
            [System.Windows.Forms.MessageBox]::Show("All installations completed!", "Complete")
        }
    })
    $timer.Start()
})