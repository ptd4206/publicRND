# ---------------------------------------------------------------------------
# Vue PACS Utility to Resart Services, Processes, IIS App Pools on Multiple Servers
# And more!
#
# Uses PowerShell GUI for easy selection
#
# ---------------------------------------------------------------------------
#
# Created with AI Assitance by Scott Mibert, Philips Healthcare September, 2025 
# Lots of input from so many awesome TCs, Support, R&D, etc!
# Feedback for improvements always welcome!
#
# ---------------------------------------------------------------------------
#
# App metadata
$AppName    = 'Restart Svc/Proc/IISAppPools (and more) on Multiple Servers'
$AppVersion = '8.5'

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Hide the PowerShell console window
Add-Type -Name Window -Namespace Console -MemberDefinition '
[DllImport("Kernel32.dll")]
public static extern IntPtr GetConsoleWindow();

[DllImport("user32.dll")]
public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);
'

$consolePtr = [Console.Window]::GetConsoleWindow()
[Console.Window]::ShowWindow($consolePtr, 0) # 0 = hide

# Load configuration from an external CSV file
$configFile = "C:\temp\ServerManagementConfig.csv"
if (-not (Test-Path $configFile)) {
    $config = [PSCustomObject]@{
        
		##REPLACE WITH YOUR SERVER HOSTNAMES
		COREServers     = 'swvpac4001,swvpac4002,swvpac4003,swvpac4004,swvpac1001,swvpac1002,swvpac6001,swvpac6002'
		DAServers       = 'swvpac4011,swvpac4012,swvpac4009,swvpac4010,swvpac1004,swvpac1005,swvpac1006,swvpac1007,swvpac1008,swvpac1009,swvpac1010,swvpac1011,swvpac1012,swvpac1013,swvpac1014,swvpac1015,swvpac6004,swvpac6005,swvpac6006,swvpac6007,swvpac6008,swvpac6009,swvpac6010,swvpac6011,swvpac6012,swvpac6013,swvpac6014,swvpac6015'
		PortalServers   = 'swvpacap4006,swvpacap4011,swvpacap4007,swvpacap4010,swvpacap1021,swvpacap1022,swvpacap1023,swvpacap1024,swvpacap1025,swvpacap1026,swvpacap1027,swvpacap1028,swvpacap1029,swvpacap1030,swvpacap6021,swvpacap6022,swvpacap6023,swvpacap6024,swvpacap6025,swvpacap6026,swvpacap6027,swvpacap6028,swvpacap6029,swvpacap6030'
		OtherServers    = 'swvpacdb1001,swvpacap1003'
        
		##ADD ADDITIONAL SERVICES TO THIS LIST (R-click Properties and use Service Name) 
		Services    = 'ElasticSearch-service,ES-Cerebro-service,Ignite,Imaginet Auto-Router Execution Module,Filebeat,FLEXlm Service, Imaginet Auto-Router Scheduling Module,Imaginet Connectivity Monitor,Imaginet DataGrid Controller,Imaginet DB Audit Server,Imaginet Failover Service,Imaginet IOCM Probe,Imaginet Loader Server,Imaginet Medilink Converter,Imaginet Medilink Listener,Imaginet Medilink Sync Listener,Imaginet MstSync Server,Imaginet MVSMain LightViewer Server,Imaginet MVSMain Secured Server,Imaginet MVSMain Server,Imaginet Non-Dicom Auto Ingestion Service,Imaginet PACS Monitor Service,Imaginet PACS Restarter Service,Imaginet Portal Restarter Service,Imaginet RisSync Server,Imaginet Startup-Shutdown,Imaginet System Check,Imaginet Task Dispatcher,Imaginet Task Scanner,Imaginet VUEEXPLORER Restarter Service,Imaginet WCF,Imaginet WFM Scheduled Queries Dispatcher,Kafka,Kafka Http Proxy,Kibana,Logstash,Metricbeat,Mirth Connect Service,MSTNormalizer,nxlog,PhilipsMemoProductExporterWrapperService,PhilipsMemoPrometheusService,ProductExporterService,Tomcat9,W3SVC,WAS,windows_exporter,Zookeeper'
        
		##ADD ADDITIONAL PROCESSES TO THIS LIST (NOTE: They are restarted with -port all)
		Processes   = 'render,svar,svarstore,svaudit,svcfg,svdds,svdidb,svdisk,svdmx,svdser,svdser_sir,svdtc,svenc,svfload,svfolder,svft,mvft,svfwd,svldr,svmigrate,svping,svqe,svreg,svrep,svris,svsecm,svsm,svspeech,svssl,svstream,svsync,svwfm,svwload'
        
		##ADD ADDITIONAL APP POOLS TO THIS LIST (R-click Basic Settings and copy the Name)
		IISAppPools = 'AdminServerAppPool,ArchiveAppPool,ArchiveAppSvcAppPool,ArchiveDataAppPool,ArchiveLoggerAppPool,ArchiveReportAppPool,ArchiveShareAppPool,ArchiveStickyNotesAppPool,ArchiveUIPatientsAppPool,ArchiveViewsAppPool,ChatAppPool,CSPublicQueryAppPool,CSPublicReportAppPool,CSPublicStickyNotesAppPool,FHIRAutoEventsServiceAppPool,FHIRDicomParserServiceAppPool,FHIRSecmServiceAppPool,FHIRServiceAppPool,FHIRWorklistItemsServiceAppPool,FHIRXDSServiceAppPool,ImageAppPool,ImportToolAppPool,LoggerAppPool,MpCFGDPAppPool,MpConfigurationAppPool,NDFAppPool,NormalizationAppPool,OrchestratorServiceAppPool,PACSAppPool,PACSAuthenticatedAppPool,PatientJacketAppPool,PortalAppPool,ReportAppPool,RISAppPool,SecmAppPool,SecmServiceProviderAppPool,SecmServiceWorkerAppPool1,SecmServiceWorkerAppPool2,SecmServiceWorkerAppPool3,SecmServiceWorkerAppPool4,SecmServiceWorkerAppPool5,SecmUsersAppPool,SessionManagerAppPool,ShareAppPool,StickyNotesAppPool,SysCfgDPAppPool,SystemConfigurationAppPool,UIPatientsAppPool,WFMOtherQueryHandlersAppPool,WFMWiManagmentAppPool'
    }
} else {
    $config = Import-Csv -Path $configFile
}

# Store the script path for editing
$global:ScriptPath = $MyInvocation.MyCommand.Path

# Global variables to store current data
$global:CurrentCoreServers = @()
$global:CurrentDAServers = @()
$global:CurrentPortalServers = @()
$global:CurrentOtherServers = @()
$global:CurrentServices = @()
$global:CurrentProcesses = @()
$global:CurrentIISAppPools = @()

# Global variable to track PowerShell windows
$global:PowerShellWindows = @()

# Initialize global variables
$global:CurrentCoreServers = $config.COREServers.Split(',')
$global:CurrentDAServers = $config.DAServers.Split(',')
$global:CurrentPortalServers = $config.PortalServers.Split(',')
$global:CurrentOtherServers = $config.OtherServers.Split(',')
$global:CurrentServices = $config.Services.Split(',') | Sort-Object
$global:CurrentProcesses = $config.Processes.Split(',')
$global:CurrentIISAppPools = $config.IISAppPools.Split(',')

# Create the main form
$form = New-Object System.Windows.Forms.Form
$form.Text = "$AppName v$AppVersion"
$form.Size = New-Object System.Drawing.Size(1350, 865)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "Sizable"

# Add form closing event to clean up PowerShell windows
$form.Add_FormClosing({
    foreach ($window in $global:PowerShellWindows) {
        try {
            if (-not $window.HasExited) {
                $window.CloseMainWindow()
            }
        } catch {
            # Window may already be closed
        }
    }
})

# ==================== ENHANCED COMMAND EXECUTION FUNCTIONS ====================
function Execute-CommandInNewWindow {
    param(
        [string]$Server,
        [string]$Command,
        [string]$WindowTitle,
        [int]$Delay = 0
    )
    
    if ($Delay -gt 0) {
        Start-Sleep -Seconds $Delay
    }
    
    # Create PowerShell script content
    $scriptContent = @"
Write-Host "Executing command on server: $Server" -ForegroundColor Green
Write-Host "Command: $Command" -ForegroundColor Yellow
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

try {
    $Command
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "Command completed successfully!" -ForegroundColor Green
} catch {
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "ERROR: `$_" -ForegroundColor Red
}

Write-Host ""
Write-Host "Press any key to close this window..." -ForegroundColor Magenta
`$null = `$Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
"@
    
    # Create temporary script file
    $tempScript = [System.IO.Path]::GetTempFileName() + ".ps1"
    Set-Content -Path $tempScript -Value $scriptContent
    
    # Start new PowerShell window
    $processInfo = New-Object System.Diagnostics.ProcessStartInfo
    $processInfo.FileName = "powershell.exe"
    $processInfo.Arguments = "-ExecutionPolicy Bypass -File `"$tempScript`""
    $processInfo.WindowStyle = "Normal"
    $processInfo.CreateNoWindow = $false
    
    # Set window title
    $process = [System.Diagnostics.Process]::Start($processInfo)
    $global:PowerShellWindows += $process
    
    # Clean up temp file after a delay (async)
    $timer = New-Object System.Timers.Timer
    $timer.Interval = 30000  # 30 seconds
    $timer.AutoReset = $false
    $timer.Add_Elapsed({
        try {
            if (Test-Path $tempScript) {
                Remove-Item $tempScript -Force -ErrorAction SilentlyContinue
            }
        } catch {}
        $timer.Dispose()
    })
    $timer.Start()
    
    return $process
}

function Execute-RebootCommand {
    param(
        [string[]]$Servers,
        [string]$ServerType
    )
    
    $result = [System.Windows.Forms.MessageBox]::Show(
        "WARNING: This will immediately REBOOT the selected $ServerType servers:`n`n$($Servers -join ', ')`n`nThis action cannot be undone and will cause downtime. Are you absolutely sure?",
        "CONFIRM SERVER REBOOT",
        "YesNo",
        "Warning"
    )
    
    if ($result -eq "Yes") {
        foreach ($server in $Servers) {
            try {
                Log-Message "REBOOTING $ServerType Server: $server"
                Restart-Computer -ComputerName $server -Force -ErrorAction Stop
                Log-Message "Reboot command sent successfully to $server"
            } catch {
                Log-Message "ERROR: Failed to reboot $server - $_"
            }
        }
    }
}

# ==================== EDIT FUNCTIONS ====================
function Show-TabbedEditWindow {
    $editForm = New-Object System.Windows.Forms.Form
    $editForm.Text = "Edit Server Management Configuration"
    $editForm.Size = New-Object System.Drawing.Size(700, 700)
    $editForm.StartPosition = "CenterParent"
    $editForm.FormBorderStyle = "FixedDialog"
    $editForm.MaximizeBox = $false
    
    # Create TabControl
    $tabControl = New-Object System.Windows.Forms.TabControl
    $tabControl.Location = New-Object System.Drawing.Point(10, 10)
    $tabControl.Size = New-Object System.Drawing.Size(665, 600)
    $editForm.Controls.Add($tabControl)
    
    # Define tab configurations
    $tabConfigs = @(
        @{Name = "CORE Servers"; Items = $global:CurrentCoreServers; Variable = "COREServers"; Sort = $false},
        @{Name = "DA Servers"; Items = $global:CurrentDAServers; Variable = "DAServers"; Sort = $false},
        @{Name = "Portal Servers"; Items = $global:CurrentPortalServers; Variable = "PortalServers"; Sort = $false},
        @{Name = "Other Servers"; Items = $global:CurrentOtherServers; Variable = "OtherServers"; Sort = $false},
        @{Name = "Services"; Items = $global:CurrentServices; Variable = "Services"; Sort = $true},
        @{Name = "Processes"; Items = $global:CurrentProcesses; Variable = "Processes"; Sort = $false},
        @{Name = "IIS App Pools"; Items = $global:CurrentIISAppPools; Variable = "IISAppPools"; Sort = $false}
    )
    
    $tabData = @{}
    
    # Create tabs
    foreach ($tabConfig in $tabConfigs) {
        $tabPage = New-Object System.Windows.Forms.TabPage
        $tabPage.Text = $tabConfig.Name
        $tabPage.UseVisualStyleBackColor = $true
        $tabControl.TabPages.Add($tabPage)
        
        # Instructions label
        $lblInstructions = New-Object System.Windows.Forms.Label
        $lblInstructions.Text = "Edit items (one per line). Click Add to add new items, Delete to remove selected items."
        $lblInstructions.Location = New-Object System.Drawing.Point(10, 10)
        $lblInstructions.Size = New-Object System.Drawing.Size(620, 30)
        $tabPage.Controls.Add($lblInstructions)
        
        # ListBox for items
        $listBox = New-Object System.Windows.Forms.ListBox
        $listBox.Location = New-Object System.Drawing.Point(10, 45)
        $listBox.Size = New-Object System.Drawing.Size(450, 400)
        $listBox.SelectionMode = "MultiExtended"
        foreach ($item in $tabConfig.Items) {
            if ($item.Trim()) {
                $listBox.Items.Add($item.Trim())
            }
        }
        $tabPage.Controls.Add($listBox)
        
        # TextBox for new/edit items
        $txtNewItem = New-Object System.Windows.Forms.TextBox
        $txtNewItem.Location = New-Object System.Drawing.Point(10, 455)
        $txtNewItem.Size = New-Object System.Drawing.Size(450, 25)
        
        # Create a label instead of using PlaceholderText
        $lblPlaceholder = New-Object System.Windows.Forms.Label
        $lblPlaceholder.Text = "Enter new item or select item above to edit:"
        $lblPlaceholder.Location = New-Object System.Drawing.Point(10, 430)
        $lblPlaceholder.Size = New-Object System.Drawing.Size(450, 20)
        $lblPlaceholder.ForeColor = [System.Drawing.Color]::Gray
        $tabPage.Controls.Add($lblPlaceholder)
        $tabPage.Controls.Add($txtNewItem)
        
        # Add button
        $btnAdd = New-Object System.Windows.Forms.Button
        $btnAdd.Text = "Add"
        $btnAdd.Location = New-Object System.Drawing.Point(470, 450)
        $btnAdd.Size = New-Object System.Drawing.Size(60, 30)
        $btnAdd.BackColor = [System.Drawing.Color]::LightGreen
        $btnAdd.Add_Click({
            if ($txtNewItem.Text.Trim() -ne "") {
                $newItem = $txtNewItem.Text.Trim()
                if ($listBox.Items -notcontains $newItem) {
                    $listBox.Items.Add($newItem)
                    $txtNewItem.Text = ""
                } else {
                    [System.Windows.Forms.MessageBox]::Show("Item already exists!", "Duplicate Item", "OK", "Warning")
                }
            }
        }.GetNewClosure())
        $tabPage.Controls.Add($btnAdd)
        
        # Update button
        $btnUpdate = New-Object System.Windows.Forms.Button
        $btnUpdate.Text = "Update"
        $btnUpdate.Location = New-Object System.Drawing.Point(470, 485)
        $btnUpdate.Size = New-Object System.Drawing.Size(60, 30)
        $btnUpdate.BackColor = [System.Drawing.Color]::LightBlue
        $btnUpdate.Add_Click({
            if ($listBox.SelectedIndex -ge 0 -and $txtNewItem.Text.Trim() -ne "") {
                $newValue = $txtNewItem.Text.Trim()
                $oldValue = $listBox.SelectedItem
                if ($listBox.Items -notcontains $newValue -or $newValue -eq $oldValue) {
                    $listBox.Items[$listBox.SelectedIndex] = $newValue
                    $txtNewItem.Text = ""
                } else {
                    [System.Windows.Forms.MessageBox]::Show("Item already exists!", "Duplicate Item", "OK", "Warning")
                }
            }
        }.GetNewClosure())
        $tabPage.Controls.Add($btnUpdate)
        
        # Delete button
        $btnDelete = New-Object System.Windows.Forms.Button
        $btnDelete.Text = "Delete"
        $btnDelete.Location = New-Object System.Drawing.Point(470, 520)
        $btnDelete.Size = New-Object System.Drawing.Size(60, 30)
        $btnDelete.BackColor = [System.Drawing.Color]::LightCoral
        $btnDelete.Add_Click({
            if ($listBox.SelectedIndices.Count -gt 0) {
                $confirm = [System.Windows.Forms.MessageBox]::Show(
                    "Are you sure you want to delete the selected item(s)?",
                    "Confirm Delete",
                    "YesNo",
                    "Question"
                )
                if ($confirm -eq "Yes") {
                    # Remove items in reverse order to maintain indices
                    $selectedIndices = $listBox.SelectedIndices | Sort-Object -Descending
                    foreach ($index in $selectedIndices) {
                        $listBox.Items.RemoveAt($index)
                    }
                }
            }
        }.GetNewClosure())
        $tabPage.Controls.Add($btnDelete)
        
        # Event handler for selecting items to edit
        $listBox.Add_Click({
            if ($listBox.SelectedIndex -ge 0) {
                $txtNewItem.Text = $listBox.SelectedItem
            }
        }.GetNewClosure())
        
        # Store tab data for later access
        $tabData[$tabConfig.Variable] = @{
            ListBox = $listBox
            Sort = $tabConfig.Sort
        }
    }
    
    # Save button
    $btnSave = New-Object System.Windows.Forms.Button
    $btnSave.Text = "Save All"
    $btnSave.Location = New-Object System.Drawing.Point(480, 620)
    $btnSave.Size = New-Object System.Drawing.Size(80, 35)
    $btnSave.BackColor = [System.Drawing.Color]::Gold
    $btnSave.Font = New-Object System.Drawing.Font("Arial", 9, [System.Drawing.FontStyle]::Bold)
    $btnSave.Add_Click({
        try {
            # Save all tabs
            foreach ($tabConfig in $tabConfigs) {
                $tabInfo = $tabData[$tabConfig.Variable]
                $listBox = $tabInfo.ListBox
                
                # Get all items from listbox
                $updatedItems = @()
                foreach ($item in $listBox.Items) {
                    if ($item.Trim()) {
                        $updatedItems += $item.Trim()
                    }
                }
                
                # Sort alphabetically if requested
                if ($tabInfo.Sort) {
                    $updatedItems = $updatedItems | Sort-Object
                }
                
                # Update the script file and global variables
                Update-ScriptFile -VariableName $tabConfig.Variable -NewValues $updatedItems
                Update-GlobalVariables -VariableName $tabConfig.Variable -NewValues $updatedItems
            }
            
            [System.Windows.Forms.MessageBox]::Show("All changes saved successfully!", "Save Complete", "OK", "Information")
            $editForm.Tag = "Saved"
            $editForm.Close()
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Error saving changes: $_", "Save Error", "OK", "Error")
        }
    })
    $editForm.Controls.Add($btnSave)
    
    # Cancel button
    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Cancel"
    $btnCancel.Location = New-Object System.Drawing.Point(570, 620)
    $btnCancel.Size = New-Object System.Drawing.Size(80, 35)
    $btnCancel.Add_Click({
        $editForm.Close()
    })
    $editForm.Controls.Add($btnCancel)
    
    # Show dialog and return result
    [void]$editForm.ShowDialog()
    return $editForm.Tag -eq "Saved"
}

function Update-ScriptFile {
    param(
        [string]$VariableName,
        [array]$NewValues
    )
    
    if (-not $global:ScriptPath -or -not (Test-Path $global:ScriptPath)) {
        throw "Script path not found. Cannot save changes to file."
    }
    
    # Read the current script content
    $scriptContent = Get-Content -Path $global:ScriptPath -Raw
    
    # Create the new value string with proper escaping
    $newValueString = ($NewValues | ForEach-Object { $_.Replace("'", "''") }) -join ','
    
    # Define the pattern to match the variable assignment - more robust regex
    $pattern = "(\s*$VariableName\s*=\s*)'([^']*(?:''[^']*)*)'(?=\s|\r|\n|$)"
    
    # Replace the old values with new ones
    $updatedContent = $scriptContent -replace $pattern, "`${1}'$newValueString'"
    
    # Write back to file
    Set-Content -Path $global:ScriptPath -Value $updatedContent -ErrorAction Stop
    
    Log-Message "Updated $VariableName in script file with $($NewValues.Count) items"
}

function Update-GlobalVariables {
    param(
        [string]$VariableName,
        [array]$NewValues
    )
    
    switch ($VariableName) {
        "COREServers" { $global:CurrentCoreServers = $NewValues }
        "DAServers" { $global:CurrentDAServers = $NewValues }
        "PortalServers" { $global:CurrentPortalServers = $NewValues }
        "OtherServers" { $global:CurrentOtherServers = $NewValues }
        "Services" { $global:CurrentServices = $NewValues }
        "Processes" { $global:CurrentProcesses = $NewValues }
        "IISAppPools" { $global:CurrentIISAppPools = $NewValues }
    }
}

function Refresh-ListBoxes {
    # Use global variables for immediate refresh instead of re-reading file
    try {
        # Clear and repopulate all listboxes with current global data
        $coreServerBox.Items.Clear()
        foreach ($server in $global:CurrentCoreServers) {
            if ($server.Trim()) { $coreServerBox.Items.Add($server.Trim()) }
        }
        
        $daServerBox.Items.Clear()
        foreach ($server in $global:CurrentDAServers) {
            if ($server.Trim()) { $daServerBox.Items.Add($server.Trim()) }
        }
        
        $portalServerBox.Items.Clear()
        foreach ($server in $global:CurrentPortalServers) {
            if ($server.Trim()) { $portalServerBox.Items.Add($server.Trim()) }
        }
        
        $otherServerBox.Items.Clear()
        foreach ($server in $global:CurrentOtherServers) {
            if ($server.Trim()) { $otherServerBox.Items.Add($server.Trim()) }
        }
        
        $serviceBox.Items.Clear()
        foreach ($service in $global:CurrentServices) {
            if ($service.Trim()) { $serviceBox.Items.Add($service.Trim()) }
        }
        
        $processBox.Items.Clear()
        foreach ($process in $global:CurrentProcesses) {
            if ($process.Trim()) { $processBox.Items.Add($process.Trim()) }
        }
        
        $poolBox.Items.Clear()
        foreach ($pool in $global:CurrentIISAppPools) {
            if ($pool.Trim()) { $poolBox.Items.Add($pool.Trim()) }
        }
        
        Log-Message "Interface refreshed successfully with updated values"
    } catch {
        Log-Message "Error refreshing interface: $_"
    }
}

# ==================== COLUMN HEADERS ====================
# Core Servers Header
$lblCoreServers = New-Object System.Windows.Forms.Label
$lblCoreServers.Text = "CORE SERVERS"
$lblCoreServers.Font = New-Object System.Drawing.Font("Arial", 9, [System.Drawing.FontStyle]::Bold)
$lblCoreServers.Location = New-Object System.Drawing.Point(10, 10)
$lblCoreServers.Size = New-Object System.Drawing.Size(120, 20)
$form.Controls.Add($lblCoreServers)

# DA Servers Header
$lblDAServers = New-Object System.Windows.Forms.Label
$lblDAServers.Text = "DA SERVERS"
$lblDAServers.Font = New-Object System.Drawing.Font("Arial", 9, [System.Drawing.FontStyle]::Bold)
$lblDAServers.Location = New-Object System.Drawing.Point(140, 10)
$lblDAServers.Size = New-Object System.Drawing.Size(120, 20)
$form.Controls.Add($lblDAServers)

# Portal Servers Header
$lblPortalServers = New-Object System.Windows.Forms.Label
$lblPortalServers.Text = "PORTAL SERVERS"
$lblPortalServers.Font = New-Object System.Drawing.Font("Arial", 9, [System.Drawing.FontStyle]::Bold)
$lblPortalServers.Location = New-Object System.Drawing.Point(270, 10)
$lblPortalServers.Size = New-Object System.Drawing.Size(120, 20)
$form.Controls.Add($lblPortalServers)

# Other Servers Header
$lblOtherServers = New-Object System.Windows.Forms.Label
$lblOtherServers.Text = "OTHER SERVERS"
$lblOtherServers.Font = New-Object System.Drawing.Font("Arial", 9, [System.Drawing.FontStyle]::Bold)
$lblOtherServers.Location = New-Object System.Drawing.Point(400, 10)
$lblOtherServers.Size = New-Object System.Drawing.Size(120, 20)
$form.Controls.Add($lblOtherServers)

# Services Header
$lblServices = New-Object System.Windows.Forms.Label
$lblServices.Text = "SERVICES:"
$lblServices.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
$lblServices.Location = New-Object System.Drawing.Point(540, 10)
$lblServices.Size = New-Object System.Drawing.Size(310, 20)
$form.Controls.Add($lblServices)

# Services Subtitle
$lblServicesSubtitle = New-Object System.Windows.Forms.Label
$lblServicesSubtitle.Text = "(restart-service)"
$lblServicesSubtitle.Font = New-Object System.Drawing.Font("Arial", 8, [System.Drawing.FontStyle]::Italic)
$lblServicesSubtitle.Location = New-Object System.Drawing.Point(540, 30)
$lblServicesSubtitle.Size = New-Object System.Drawing.Size(310, 20)
$form.Controls.Add($lblServicesSubtitle)

# Processes Header
$lblProcesses = New-Object System.Windows.Forms.Label
$lblProcesses.Text = "PROCESSES:"
$lblProcesses.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
$lblProcesses.Location = New-Object System.Drawing.Point(870, 10)
$lblProcesses.Size = New-Object System.Drawing.Size(180, 20)
$form.Controls.Add($lblProcesses)

# Processes Subtitle
$lblProcessesSubtitle = New-Object System.Windows.Forms.Label
$lblProcessesSubtitle.Text = "(admincommand -p all)"
$lblProcessesSubtitle.Font = New-Object System.Drawing.Font("Arial", 8, [System.Drawing.FontStyle]::Italic)
$lblProcessesSubtitle.Location = New-Object System.Drawing.Point(870, 30)
$lblProcessesSubtitle.Size = New-Object System.Drawing.Size(180, 20)
$form.Controls.Add($lblProcessesSubtitle)

# IIS App Pools Header
$lblIISAppPools = New-Object System.Windows.Forms.Label
$lblIISAppPools.Text = "IISAppPools"
$lblIISAppPools.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
$lblIISAppPools.Location = New-Object System.Drawing.Point(1070, 10)
$lblIISAppPools.Size = New-Object System.Drawing.Size(250, 20)
$form.Controls.Add($lblIISAppPools)

# IIS App Pools Subtitle
$lblIISSubtitle = New-Object System.Windows.Forms.Label
$lblIISSubtitle.Text = "(Restart-WebAppPool)"
$lblIISSubtitle.Font = New-Object System.Drawing.Font("Arial", 8, [System.Drawing.FontStyle]::Italic)
$lblIISSubtitle.Location = New-Object System.Drawing.Point(1070, 30)
$lblIISSubtitle.Size = New-Object System.Drawing.Size(250, 20)
$form.Controls.Add($lblIISSubtitle)

# ==================== CHECKEDLISTBOXES ====================
$tooltip = New-Object System.Windows.Forms.ToolTip

# Core Servers Box (Light Blue Background)
$coreServerBox = New-Object System.Windows.Forms.CheckedListBox
$coreServerBox.Location = New-Object System.Drawing.Point(10, 55)
$coreServerBox.Size = New-Object System.Drawing.Size(120, 485)
$coreServerBox.CheckOnClick = $true
$coreServerBox.BackColor = [System.Drawing.Color]::FromArgb(222, 235, 246)
foreach ($server in $global:CurrentCoreServers) {
    if ($server.Trim()) { $coreServerBox.Items.Add($server.Trim()) }
}
$coreServerBox.Anchor = "Top, Left"
$tooltip.SetToolTip($coreServerBox, "Select CORE servers to manage.")
$form.Controls.Add($coreServerBox)

# DA Servers Box (Light Blue Background)
$daServerBox = New-Object System.Windows.Forms.CheckedListBox
$daServerBox.Location = New-Object System.Drawing.Point(140, 55)
$daServerBox.Size = New-Object System.Drawing.Size(120, 485)
$daServerBox.CheckOnClick = $true
$daServerBox.BackColor = [System.Drawing.Color]::FromArgb(222, 235, 246)
foreach ($server in $global:CurrentDAServers) {
    if ($server.Trim()) { $daServerBox.Items.Add($server.Trim()) }
}
$daServerBox.Anchor = "Top, Left"
$tooltip.SetToolTip($daServerBox, "Select DA servers to manage.")
$form.Controls.Add($daServerBox)

# Portal Servers Box (Light Blue Background)
$portalServerBox = New-Object System.Windows.Forms.CheckedListBox
$portalServerBox.Location = New-Object System.Drawing.Point(270, 55)
$portalServerBox.Size = New-Object System.Drawing.Size(120, 485)
$portalServerBox.CheckOnClick = $true
$portalServerBox.BackColor = [System.Drawing.Color]::FromArgb(222, 235, 246)
foreach ($server in $global:CurrentPortalServers) {
    if ($server.Trim()) { $portalServerBox.Items.Add($server.Trim()) }
}
$portalServerBox.Anchor = "Top, Left"
$tooltip.SetToolTip($portalServerBox, "Select PORTAL servers to manage.")
$form.Controls.Add($portalServerBox)

# Other Servers Box (Light Blue Background)
$otherServerBox = New-Object System.Windows.Forms.CheckedListBox
$otherServerBox.Location = New-Object System.Drawing.Point(400, 55)
$otherServerBox.Size = New-Object System.Drawing.Size(120, 485)
$otherServerBox.CheckOnClick = $true
$otherServerBox.BackColor = [System.Drawing.Color]::FromArgb(222, 235, 246)
foreach ($server in $global:CurrentOtherServers) {
    if ($server.Trim()) { $otherServerBox.Items.Add($server.Trim()) }
}
$otherServerBox.Anchor = "Top, Left"
$tooltip.SetToolTip($otherServerBox, "Select OTHER servers to manage.")
$form.Controls.Add($otherServerBox)

# Services Box (Light Green Background)
$serviceBox = New-Object System.Windows.Forms.CheckedListBox
$serviceBox.Location = New-Object System.Drawing.Point(540, 55)
$serviceBox.Size = New-Object System.Drawing.Size(310, 485)
$serviceBox.CheckOnClick = $true
$serviceBox.BackColor = [System.Drawing.Color]::FromArgb(226, 239, 217)
foreach ($service in $global:CurrentServices) {
    if ($service.Trim()) { $serviceBox.Items.Add($service.Trim()) }
}
$serviceBox.Anchor = "Top, Left"
$tooltip.SetToolTip($serviceBox, "Select services to manage.")
$form.Controls.Add($serviceBox)

# Processes Box (Light Orange Background)
$processBox = New-Object System.Windows.Forms.CheckedListBox
$processBox.Location = New-Object System.Drawing.Point(870, 55)
$processBox.Size = New-Object System.Drawing.Size(180, 485)
$processBox.CheckOnClick = $true
$processBox.BackColor = [System.Drawing.Color]::FromArgb(251, 229, 213)
foreach ($process in $global:CurrentProcesses) {
    if ($process.Trim()) { $processBox.Items.Add($process.Trim()) }
}
$processBox.Anchor = "Top, Left"
$tooltip.SetToolTip($processBox, "Select processes to restart.")
$form.Controls.Add($processBox)

# IIS App Pools Box (Light Yellow Background)
$poolBox = New-Object System.Windows.Forms.CheckedListBox
$poolBox.Location = New-Object System.Drawing.Point(1070, 55)
$poolBox.Size = New-Object System.Drawing.Size(250, 485)
$poolBox.CheckOnClick = $true
$poolBox.BackColor = [System.Drawing.Color]::FromArgb(255, 242, 204)
foreach ($pool in $global:CurrentIISAppPools) {
    if ($pool.Trim()) { $poolBox.Items.Add($pool.Trim()) }
}
$poolBox.Anchor = "Top, Left"
$tooltip.SetToolTip($poolBox, "Select IIS App Pools to restart.")
$form.Controls.Add($poolBox)

# Core Servers Task Dropdown
$coreTaskDropdown = New-Object System.Windows.Forms.ComboBox
$coreTaskDropdown.Location = New-Object System.Drawing.Point(10, 545)
$coreTaskDropdown.Size = New-Object System.Drawing.Size(120, 25)
$coreTaskDropdown.DropDownStyle = "DropDownList"
$coreTaskDropdown.Items.AddRange(@(
    "--select command--",
    "STOP APP",
    "START APP",
    "STOP APP/DB",
    "START APP/DB",
    "REBOOT Server(s)"
))
$coreTaskDropdown.SelectedIndex = 0
$form.Controls.Add($coreTaskDropdown)

# Core Execute Button (RED with Light Grey Text)
$btnCoreExecute = New-Object System.Windows.Forms.Button
$btnCoreExecute.Text = "EXECUTE"
$btnCoreExecute.Location = New-Object System.Drawing.Point(10, 575)
$btnCoreExecute.Size = New-Object System.Drawing.Size(120, 25)
$btnCoreExecute.BackColor = [System.Drawing.Color]::FromArgb(245, 65, 65)
$btnCoreExecute.ForeColor = [System.Drawing.Color]::LightGray
$btnCoreExecute.Font = New-Object System.Drawing.Font("Arial", 9, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($btnCoreExecute)

# DA Servers Task Dropdown
$daTaskDropdown = New-Object System.Windows.Forms.ComboBox
$daTaskDropdown.Location = New-Object System.Drawing.Point(140, 545)
$daTaskDropdown.Size = New-Object System.Drawing.Size(120, 25)
$daTaskDropdown.DropDownStyle = "DropDownList"
$daTaskDropdown.Items.AddRange(@(
    "--select command--",
    "STOP APP",
    "START APP",
    "REBOOT Server(s)"
))
$daTaskDropdown.SelectedIndex = 0
$form.Controls.Add($daTaskDropdown)

# DA Execute Button (RED with Light Grey Text)
$btnDAExecute = New-Object System.Windows.Forms.Button
$btnDAExecute.Text = "EXECUTE"
$btnDAExecute.Location = New-Object System.Drawing.Point(140, 575)
$btnDAExecute.Size = New-Object System.Drawing.Size(120, 25)
$btnDAExecute.BackColor = [System.Drawing.Color]::FromArgb(245, 65, 65)
$btnDAExecute.ForeColor = [System.Drawing.Color]::LightGray
$btnDAExecute.Font = New-Object System.Drawing.Font("Arial", 9, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($btnDAExecute)

# Portal Servers Task Dropdown
$portalTaskDropdown = New-Object System.Windows.Forms.ComboBox
$portalTaskDropdown.Location = New-Object System.Drawing.Point(270, 545)
$portalTaskDropdown.Size = New-Object System.Drawing.Size(120, 25)
$portalTaskDropdown.DropDownStyle = "DropDownList"
$portalTaskDropdown.Items.AddRange(@(
    "--select command--",
    "Portal STOP",
    "Portal START",
    "Portal RESTART",
    "REBOOT Server(s)"
))
$portalTaskDropdown.SelectedIndex = 0
$form.Controls.Add($portalTaskDropdown)

# Portal Execute Button (RED with Light Grey Text)
$btnPortalExecute = New-Object System.Windows.Forms.Button
$btnPortalExecute.Text = "EXECUTE"
$btnPortalExecute.Location = New-Object System.Drawing.Point(270, 575)
$btnPortalExecute.Size = New-Object System.Drawing.Size(120, 25)
$btnPortalExecute.BackColor = [System.Drawing.Color]::FromArgb(245, 65, 65)
$btnPortalExecute.ForeColor = [System.Drawing.Color]::LightGray
$btnPortalExecute.Font = New-Object System.Drawing.Font("Arial", 9, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($btnPortalExecute)

# Other Servers Task Dropdown
$otherTaskDropdown = New-Object System.Windows.Forms.ComboBox
$otherTaskDropdown.Location = New-Object System.Drawing.Point(400, 545)
$otherTaskDropdown.Size = New-Object System.Drawing.Size(120, 25)
$otherTaskDropdown.DropDownStyle = "DropDownList"
$otherTaskDropdown.Items.AddRange(@(
    "--select command--",
    "ES RESTART",
    "EA STOP",
    "EA START",
    "EA RESTART",
    "EA STATUS",
    "RA STOP",
    "RA START",
    "RA RESTART",
    "REBOOT Server(s)"
))
$otherTaskDropdown.SelectedIndex = 0
$form.Controls.Add($otherTaskDropdown)

# Other Execute Button (RED with Light Grey Text)
$btnOtherExecute = New-Object System.Windows.Forms.Button
$btnOtherExecute.Text = "EXECUTE"
$btnOtherExecute.Location = New-Object System.Drawing.Point(400, 575)
$btnOtherExecute.Size = New-Object System.Drawing.Size(120, 25)
$btnOtherExecute.BackColor = [System.Drawing.Color]::FromArgb(245, 65, 65)
$btnOtherExecute.ForeColor = [System.Drawing.Color]::LightGray
$btnOtherExecute.Font = New-Object System.Drawing.Font("Arial", 9, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($btnOtherExecute)

# ==================== SECTION BUTTONS ====================
# Services Section Buttons - Corrected Layout aligned with EXECUTE buttons
$btnServiceRestart = New-Object System.Windows.Forms.Button
$btnServiceRestart.Text = "RESTART"
$btnServiceRestart.Location = New-Object System.Drawing.Point(680, 550)
$btnServiceRestart.Size = New-Object System.Drawing.Size(75, 25)
$btnServiceRestart.BackColor = [System.Drawing.Color]::LightBlue
$form.Controls.Add($btnServiceRestart)

# START Button (GREEN with Light Grey Text)
$btnServiceStart = New-Object System.Windows.Forms.Button
$btnServiceStart.Text = "START"
$btnServiceStart.Location = New-Object System.Drawing.Point(540, 580)
$btnServiceStart.Size = New-Object System.Drawing.Size(55, 25)
$btnServiceStart.BackColor = [System.Drawing.Color]::FromArgb(168, 208, 141)
$form.Controls.Add($btnServiceStart)

# STOP Button (RED with Light Grey Text)
$btnServiceStop = New-Object System.Windows.Forms.Button
$btnServiceStop.Text = "STOP"
$btnServiceStop.Location = New-Object System.Drawing.Point(540, 550)
$btnServiceStop.Size = New-Object System.Drawing.Size(55, 25)
$btnServiceStop.BackColor = [System.Drawing.Color]::LightCoral
$form.Controls.Add($btnServiceStop)

# STATUS Button (Grey with Blue text)
$btnServiceEnable = New-Object System.Windows.Forms.Button
$btnServiceEnable.Text = "ENABLE"
$btnServiceEnable.Location = New-Object System.Drawing.Point(600, 580)
$btnServiceEnable.Size = New-Object System.Drawing.Size(65, 25)
$btnServiceEnable.BackColor = [System.Drawing.Color]::LightGray
$btnServiceEnable.ForeColor = [System.Drawing.Color]::Green
$btnServiceEnable.Font = New-Object System.Drawing.Font("Arial", 8, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($btnServiceEnable)

# DISABLE Button (Grey with Red text)
$btnServiceDisable = New-Object System.Windows.Forms.Button
$btnServiceDisable.Text = "DISABLE"
$btnServiceDisable.Location = New-Object System.Drawing.Point(600, 550)
$btnServiceDisable.Size = New-Object System.Drawing.Size(65, 25)
$btnServiceDisable.BackColor = [System.Drawing.Color]::LightGray
$btnServiceDisable.ForeColor = [System.Drawing.Color]::Red
$btnServiceDisable.Font = New-Object System.Drawing.Font("Arial", 8, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($btnServiceDisable)

# GET INFO Button (Grey with Blue text like STATUS button)
$btnServiceGetInfo = New-Object System.Windows.Forms.Button
$btnServiceGetInfo.Text = "GET INFO"
$btnServiceGetInfo.Location = New-Object System.Drawing.Point(765, 580)
$btnServiceGetInfo.Size = New-Object System.Drawing.Size(70, 25)
$btnServiceGetInfo.BackColor = [System.Drawing.Color]::LightGray
$btnServiceGetInfo.ForeColor = [System.Drawing.Color]::Blue
$btnServiceGetInfo.Font = New-Object System.Drawing.Font("Arial", 8, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($btnServiceGetInfo)

# TASKKILL Button (RED with Light Grey Text)
$btnServiceTaskKill = New-Object System.Windows.Forms.Button
$btnServiceTaskKill.Text = "TASKKILL"
$btnServiceTaskKill.Location = New-Object System.Drawing.Point(765, 550)
$btnServiceTaskKill.Size = New-Object System.Drawing.Size(70, 25)
$btnServiceTaskKill.BackColor = [System.Drawing.Color]::FromArgb(245, 65, 65)
$btnServiceTaskKill.ForeColor = [System.Drawing.Color]::LightGray
$form.Controls.Add($btnServiceTaskKill)

# Processes Section Buttons
$btnProcessRestart = New-Object System.Windows.Forms.Button
$btnProcessRestart.Text = "RESTART_SVC"
$btnProcessRestart.Location = New-Object System.Drawing.Point(870, 550)
$btnProcessRestart.Size = New-Object System.Drawing.Size(120, 25)
$btnProcessRestart.BackColor = [System.Drawing.Color]::LightBlue
$form.Controls.Add($btnProcessRestart)


# LS | find Button (Grey with Purple text)
$btnProcessLS = New-Object System.Windows.Forms.Button
$btnProcessLS.Text = "LS |find"
$btnProcessLS.Location = New-Object System.Drawing.Point(870, 580)
$btnProcessLS.Size = New-Object System.Drawing.Size(65, 25)
$btnProcessLS.BackColor = [System.Drawing.Color]::LightGray
$btnProcessLS.ForeColor = [System.Drawing.Color]::Purple
$btnProcessLS.Font = New-Object System.Drawing.Font("Arial", 8, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($btnProcessLS)

# IIS App Pools Section Buttons
$btnIISRecycle = New-Object System.Windows.Forms.Button
$btnIISRecycle.Text = "RECYCLE"
$btnIISRecycle.Location = New-Object System.Drawing.Point(1165, 550)
$btnIISRecycle.Size = New-Object System.Drawing.Size(75, 25)
$btnIISRecycle.BackColor = [System.Drawing.Color]::LightBlue
$form.Controls.Add($btnIISRecycle)

# START Button (GREEN with Light Grey Text)
$btnIISStart = New-Object System.Windows.Forms.Button
$btnIISStart.Text = "START"
$btnIISStart.Location = New-Object System.Drawing.Point(1100, 580)
$btnIISStart.Size = New-Object System.Drawing.Size(55, 25)
$btnIISStart.BackColor = [System.Drawing.Color]::FromArgb(168, 208, 141)
$form.Controls.Add($btnIISStart)

# STOP Button (RED with Light Grey Text)
$btnIISStop = New-Object System.Windows.Forms.Button
$btnIISStop.Text = "STOP"
$btnIISStop.Location = New-Object System.Drawing.Point(1100, 550)
$btnIISStop.Size = New-Object System.Drawing.Size(55, 25)
$btnIISStop.BackColor = [System.Drawing.Color]::LightCoral
$form.Controls.Add($btnIISStop)

# IIS Recycle Settings Button - IMPROVED VERSION
$btnIISRecycleSettings = New-Object System.Windows.Forms.Button
$btnIISRecycleSettings.Text = "Recycle Settings"
$btnIISRecycleSettings.Location = New-Object System.Drawing.Point(1165, 580)
$btnIISRecycleSettings.Size = New-Object System.Drawing.Size(75, 30)
$btnIISRecycleSettings.BackColor = [System.Drawing.Color]::LightGray
$btnIISRecycleSettings.ForeColor = [System.Drawing.Color]::Blue
$btnIISRecycleSettings.Font = New-Object System.Drawing.Font("Arial", 7, [System.Drawing.FontStyle]::Bold)
$btnIISRecycleSettings.Visible = $false
$btnIISRecycleSettings.Enabled = $false
$btnIISRecycleSettings.Add_Click({
    if (-not (Validate-Servers)) { return }

    $selectedServers = Get-AllSelectedServers
    Log-Message "=== IIS RECYCLE SETTINGS REPORT OPERATION STARTED ==="
    Log-Message "Selected Servers: $($selectedServers -join ', ')"

    # Collect IIS App Pool settings per server
    $allAppPoolData = @{}
    $appPoolNameSet = New-Object System.Collections.Generic.HashSet[string]

    foreach ($server in $selectedServers | Sort-Object) {
        try {
            $res = Invoke-Command -ComputerName $server -ScriptBlock {
                $serverName = ${env:COMPUTERNAME}
                try {
                    Import-Module WebAdministration -ErrorAction Stop
                } catch {
                    return [pscustomobject]@{ Server = $serverName; Error = "WebAdministration module not available" }
                }
                try {
                    $pools = Get-ChildItem IIS:\AppPools | ForEach-Object {
                        $timeHours = 0
                        if ($_.recycling.periodicRestart.time) { $timeHours = $_.recycling.periodicRestart.time.TotalHours }
                        $times = ''
                        try {
                            $times = ($_.recycling.periodicRestart.schedule.collection | ForEach-Object { $_.value }) -join ','
                        } catch { $times = '' }
                        [PSCustomObject]@{
                            Name            = $_.Name
                            PrivateMemoryKB = $_.recycling.periodicRestart.privateMemory
                            Requests        = $_.recycling.periodicRestart.requests
                            SpecificTime    = $times
                            TimeHours       = $timeHours
                        }
                    }
                } catch {
                    return [pscustomobject]@{ Server = $serverName; Error = "Failed to enumerate AppPools: $($_.Exception.Message)" }
                }
                return [pscustomobject]@{ Server = $serverName; Pools = $pools }
            } -ErrorAction Stop

            $allAppPoolData[$server] = @{}
            if ($null -ne $res -and $null -ne $res.Error -and ($res.Error -ne "")) {
                Log-Message "IIS data error on ${server}: $($res.Error)"
            } elseif ($null -ne $res -and $null -ne $res.Pools) {
                $poolCount = 0
                foreach ($p in $res.Pools) {
                    $name = [string]$p.Name
                    if (-not [string]::IsNullOrWhiteSpace($name)) {
                        [void]$appPoolNameSet.Add($name)
                        $mem = [long]0
                        [long]::TryParse([string]$p.PrivateMemoryKB, [ref]$mem) | Out-Null
                        $req = [int]0
                        [int]::TryParse([string]$p.Requests, [ref]$req) | Out-Null
                        $hrs = [double]0
                        [double]::TryParse([string]$p.TimeHours, [ref]$hrs) | Out-Null
                        $allAppPoolData[$server][$name] = [PSCustomObject]@{
                            AppPoolName     = $name
                            PrivateMemoryKB = $mem
                            Requests        = $req
                            SpecificTime    = [string]$p.SpecificTime
                            TimeHours       = $hrs
                        }
                        $poolCount++
                    }
                }
                Log-Message "Retrieved $poolCount app pool(s) from ${server}"
            } else {
                Log-Message "No Pools or Error returned from ${server}"
            }
        } catch {
            Log-Message "ERROR retrieving IIS settings from $server - $_"
            $allAppPoolData[$server] = @{}
        }
    }

    # Build consolidated, unique app pool name list from collected data (robust across PS versions)
    $allAppPoolNames = @()
    foreach ($server in $selectedServers | Sort-Object) {
        if ($allAppPoolData.ContainsKey($server) -and $allAppPoolData[$server].Count -gt 0) {
            $allAppPoolNames += $allAppPoolData[$server].Keys
        }
    }
    $allAppPoolNames = $allAppPoolNames | Sort-Object -Unique
    $report = @()

    # Consolidated report header 
    $report += "=== IIS APP POOL RECYCLE SETTINGS REPORT ==="
    $report += "Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    $report += "Servers: $($selectedServers.Count) selected"
    $report += "Total Unique App Pools Found: $($allAppPoolNames.Count)"
    $report += ""

    # TABLE 1: PRIVATE MEMORY LIMIT
    $report += "=== TABLE 1: PRIVATE MEMORY LIMIT (KB) ==="
    $report += ""
    $headerRow = "AppPoolName".PadRight(30)
    foreach ($server in $selectedServers | Sort-Object) { $headerRow += $server.PadLeft(15) }
    $report += $headerRow
    $report += (New-Object string('-', $headerRow.Length))
    foreach ($appPoolName in $allAppPoolNames) {
        $dataRow = $appPoolName.PadRight(30)
        foreach ($server in $selectedServers | Sort-Object) {
            $value = "N/A"
            if ($allAppPoolData.ContainsKey($server) -and $allAppPoolData[$server].ContainsKey($appPoolName)) {
                $memRaw = $allAppPoolData[$server][$appPoolName].PrivateMemoryKB
                $mem = 0
                [int64]::TryParse([string]$memRaw, [ref]$mem) | Out-Null
                $value = if ($mem -gt 0) { "{0:N0}" -f $mem } else { "0" }
            }
            $dataRow += $value.PadLeft(15)
        }
        $report += $dataRow
    }
    $report += ""
    $report += ""

    # TABLE 2: REQUEST LIMITS
    $report += "=== TABLE 2: REQUEST LIMITS ==="
    $report += ""
    $headerRow = "AppPoolName".PadRight(30)
    foreach ($server in $selectedServers | Sort-Object) { $headerRow += $server.PadLeft(15) }
    $report += $headerRow
    $report += (New-Object string('-', $headerRow.Length))
    foreach ($appPoolName in $allAppPoolNames) {
        $dataRow = $appPoolName.PadRight(30)
        foreach ($server in $selectedServers | Sort-Object) {
            $value = "N/A"
            if ($allAppPoolData.ContainsKey($server) -and $allAppPoolData[$server].ContainsKey($appPoolName)) {
                $reqRaw = $allAppPoolData[$server][$appPoolName].Requests
                $req = 0
                [int]::TryParse([string]$reqRaw, [ref]$req) | Out-Null
                $value = if ($req -gt 0) { "{0:N0}" -f $req } else { "0" }
            }
            $dataRow += $value.PadLeft(15)
        }
        $report += $dataRow
    }
    $report += ""
    $report += ""

    # TABLE 3: SPECIFIC RECYCLE TIMES
    $report += "=== TABLE 3: SPECIFIC RECYCLE TIMES ==="
    $report += ""
    $headerRow = "AppPoolName".PadRight(30)
    foreach ($server in $selectedServers | Sort-Object) { $headerRow += $server.PadLeft(15) }
    $report += $headerRow
    $report += (New-Object string('-', $headerRow.Length))
    foreach ($appPoolName in $allAppPoolNames) {
        $dataRow = $appPoolName.PadRight(30)
        foreach ($server in $selectedServers | Sort-Object) {
            $value = "N/A"
            if ($allAppPoolData.ContainsKey($server) -and $allAppPoolData[$server].ContainsKey($appPoolName)) {
                $timeValue = $allAppPoolData[$server][$appPoolName].SpecificTime
                if ([string]::IsNullOrWhiteSpace($timeValue)) { $value = "None" } else { $value = $timeValue }
            }
            $dataRow += $value.PadLeft(15)
        }
        $report += $dataRow
    }
    $report += ""
    $report += ""

    # TABLE 4: TIME-BASED RECYCLE (HOURS)
    $report += "=== TABLE 4: TIME-BASED RECYCLE (HOURS) ==="
    $report += ""
    $headerRow = "AppPoolName".PadRight(30)
    foreach ($server in $selectedServers | Sort-Object) { $headerRow += $server.PadLeft(15) }
    $report += $headerRow
    $report += (New-Object string('-', $headerRow.Length))
    foreach ($appPoolName in $allAppPoolNames) {
        $dataRow = $appPoolName.PadRight(30)
        foreach ($server in $selectedServers | Sort-Object) {
            $value = "N/A"
            if ($allAppPoolData.ContainsKey($server) -and $allAppPoolData[$server].ContainsKey($appPoolName)) {
                $hrsRaw = $allAppPoolData[$server][$appPoolName].TimeHours
                $hrs = 0.0
                [double]::TryParse([string]$hrsRaw, [ref]$hrs) | Out-Null
                $value = if ($hrs -gt 0) { $hrs.ToString("F1") } else { "0" }
            }
            $dataRow += $value.PadLeft(15)
        }
        $report += $dataRow
    }
    $report += ""
    $report += ""

    # DETAILED RESULTS BY SERVER
    $report += "=== DETAILED RESULTS BY SERVER ==="
    $report += ""
    foreach ($server in $selectedServers | Sort-Object) {
        $report += "===== Results for ${server} ====="
        if ($allAppPoolData.ContainsKey($server) -and $allAppPoolData[$server].Count -gt 0) {
            $report += ""
            $report += "AppPoolName".PadRight(35) + "PrivateMemoryKB".PadLeft(15) + "Requests".PadLeft(12) + "SpecificTime".PadLeft(15) + "TimeHours".PadLeft(12)
            $report += (New-Object string('-', 89))
            $serverAppPools = $allAppPoolData[$server].Keys | Sort-Object
            foreach ($appPoolName in $serverAppPools) {
                $appPool = $allAppPoolData[$server][$appPoolName]
                $memVal = 0; [int64]::TryParse([string]$appPool.PrivateMemoryKB, [ref]$memVal) | Out-Null
                $reqVal = 0; [int]::TryParse([string]$appPool.Requests, [ref]$reqVal) | Out-Null
                $hrsVal = 0.0; [double]::TryParse([string]$appPool.TimeHours, [ref]$hrsVal) | Out-Null
                $memoryDisplay = if ($memVal -gt 0) { "{0:N0}" -f $memVal } else { "0" }
                $requestsDisplay = if ($reqVal -gt 0) { "{0:N0}" -f $reqVal } else { "0" }
                $timeDisplay = if ([string]::IsNullOrWhiteSpace([string]$appPool.SpecificTime)) { "None" } else { [string]$appPool.SpecificTime }
                $hoursDisplay = if ($hrsVal -gt 0) { $hrsVal.ToString("F1") } else { "0" }
                $line = $appPool.AppPoolName.PadRight(35) + $memoryDisplay.PadLeft(15) + $requestsDisplay.PadLeft(12) + $timeDisplay.PadLeft(15) + $hoursDisplay.PadLeft(12)
                $report += $line
            }
        } else {
            $report += "No data retrieved or connection failed (check errors in log)"
        }
        $report += ""
        $report += ""
    }

    # Create results window
    $resultsForm = New-Object System.Windows.Forms.Form
    $resultsForm.Text = "IIS App Pool Recycle Settings Report - $($selectedServers.Count) Servers"
    $resultsForm.Size = New-Object System.Drawing.Size(1500, 700)
    $resultsForm.StartPosition = "CenterParent"
    $resultsForm.FormBorderStyle = "Sizable"
    # Allow ESC to close this task window
    $resultsForm.KeyPreview = $true
    $resultsForm.Add_KeyDown({ if ($_.KeyCode -eq 'Escape') { $resultsForm.Close() } })

    # Results TextBox
    $resultsTextBox = New-Object System.Windows.Forms.TextBox
    $resultsTextBox.Location = New-Object System.Drawing.Point(10, 10)
    $resultsTextBox.Size = New-Object System.Drawing.Size(1460, 600)
    $resultsTextBox.Multiline = $true
    $resultsTextBox.ScrollBars = "Vertical"
    $resultsTextBox.Font = New-Object System.Drawing.Font("Consolas", 9)
    $resultsTextBox.WordWrap = $false
    $resultsTextBox.ReadOnly = $true
    $resultsTextBox.Text = $report -join "`r`n"
    $resultsTextBox.SelectionStart = 0
    $resultsTextBox.SelectionLength = 0
    $resultsTextBox.Anchor = "Top, Left, Bottom, Right"
    $resultsForm.Controls.Add($resultsTextBox)

    # Copy to Clipboard Button
    $copyButton = New-Object System.Windows.Forms.Button
    $copyButton.Text = "Copy to Clipboard"
    $copyButton.Location = New-Object System.Drawing.Point(450, 620)
    $copyButton.Size = New-Object System.Drawing.Size(120, 30)
    $copyButton.Anchor = "Bottom"
    $copyButton.BackColor = [System.Drawing.Color]::LightGreen
    $copyButton.Font = New-Object System.Drawing.Font("Arial", 9, [System.Drawing.FontStyle]::Bold)
    $copyButton.Add_Click({
        try {
            [System.Windows.Forms.Clipboard]::SetText($resultsTextBox.Text)
            $copyButton.Text = "Copied!"
            $copyButton.BackColor = [System.Drawing.Color]::Gold
            $timer = New-Object System.Windows.Forms.Timer
            $timer.Interval = 1500
            $timer.Add_Tick({
                $copyButton.Text = "Copy to Clipboard"
                $copyButton.BackColor = [System.Drawing.Color]::LightGreen
                $timer.Stop()
                $timer.Dispose()
            })
            $timer.Start()
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Failed to copy to clipboard: $_", "Copy Error", "OK", "Error")
        }
    })
    $resultsForm.Controls.Add($copyButton)

    # Close Button
    $closeButton = New-Object System.Windows.Forms.Button
    $closeButton.Text = "Close"
    $closeButton.Location = New-Object System.Drawing.Point(580, 620)
    $closeButton.Size = New-Object System.Drawing.Size(80, 30)
    $closeButton.Anchor = "Bottom"
    $closeButton.Add_Click({ $resultsForm.Close() })
    $resultsForm.Controls.Add($closeButton)

    [void]$resultsForm.ShowDialog()

    Log-Message "=== IIS RECYCLE SETTINGS REPORT OPERATION COMPLETED ==="
})
$form.Controls.Add($btnIISRecycleSettings)


# ==================== LOG TEXTBOX ====================
$logBox = New-Object System.Windows.Forms.RichTextBox
$logBox.Location = New-Object System.Drawing.Point(10, 615)
$logBox.Size = New-Object System.Drawing.Size(1310, 144)
$logBox.Multiline = $true
$logBox.ScrollBars = [System.Windows.Forms.RichTextBoxScrollBars]::Vertical
$logBox.Font = New-Object System.Drawing.Font("Consolas", 9)
$logBox.WordWrap = $true
$logBox.Anchor = "Top, Left, Bottom, Right"
$form.Controls.Add($logBox)

# ==================== LOGGING FUNCTION ====================
function Log-Message {
    param (
        [string]$Message
    )
    $logFile = "C:\temp\RESTART_SVC_$(Get-Date -Format 'yyyyMMdd').log"
    $timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    $logMessage = "$timestamp - $Message"
    if (-not (Test-Path "C:\temp")) { New-Item -Path "C:\temp" -ItemType Directory -Force }
    Add-Content -Path $logFile -Value $logMessage -ErrorAction SilentlyContinue
    # Highlight negative/failed responses in bold red (case-insensitive)
    # Avoid highlighting normal operational words like 'Stopped'. Build regex from terms to avoid embedded newlines.
    $negTerms = @(
        'fail(ed|ure|ing)?',
        'error',
        'not\s+found',
        'could\s+not',
        'cannot|can''t',
        'denied|refused',
        'exception',
        'time\s*out|timed\s*out',
        'unreachable',
        'invalid|missing',
        'unavailable|not\s+available|results?\s+not\s+available',
        'could\s+not\s+start|failed\s+to\s+start',
        'could\s+not\s+stop|failed\s+to\s+stop',
        'failed\s+to\s+recycle|could\s+not\s+recycle|recycle\s+failed',
        'failed\s+to\s+restart|could\s+not\s+restart|restart\s+unsuccessful',
        'failed\s+to\s+connect|could\s+not\s+connect|connection\s+failed',
        'not\s+running',
        'service\s+not\s+found',
        'app\s*pool\s+not\s+found',
        'restart\s+failed',
        'no\s+data'
    )
    $negPattern = '(' + ($negTerms -join '|') + ')'
    $isNegative = ($Message -match $negPattern)

    # Move caret to end and set selection formatting
    $logBox.SelectionStart = $logBox.TextLength
    $logBox.SelectionLength = 0
    if ($isNegative) {
        $logBox.SelectionColor = [System.Drawing.Color]::Red
        $logBox.SelectionFont  = New-Object System.Drawing.Font($logBox.Font, [System.Drawing.FontStyle]::Bold)
    } else {
        $logBox.SelectionColor = [System.Drawing.Color]::Black
        $logBox.SelectionFont  = New-Object System.Drawing.Font($logBox.Font, [System.Drawing.FontStyle]::Regular)
    }
    $logBox.SelectedText = "$logMessage`n"

    # Reset selection and scroll
    $logBox.SelectionStart = $logBox.Text.Length
    $logBox.SelectionLength = 0
    $logBox.ScrollToCaret()
}

# ==================== VALIDATION FUNCTION ====================
function Validate-Servers {
    $allSelectedServers = @()
    $allSelectedServers += $coreServerBox.CheckedItems
    $allSelectedServers += $daServerBox.CheckedItems
    $allSelectedServers += $portalServerBox.CheckedItems
    $allSelectedServers += $otherServerBox.CheckedItems
    
    if ($allSelectedServers.Count -eq 0) {
        Log-Message "ERROR: No servers selected. Please select at least one server."
        [System.Windows.Forms.MessageBox]::Show("Please select at least one server before proceeding.", "No Servers Selected", "OK", "Warning")
        return $false
    }
    return $true
}

# ==================== GET ALL SELECTED SERVERS FUNCTION ====================
function Get-AllSelectedServers {
    $allSelectedServers = @()
    $allSelectedServers += $coreServerBox.CheckedItems
    $allSelectedServers += $daServerBox.CheckedItems
    $allSelectedServers += $portalServerBox.CheckedItems
    $allSelectedServers += $otherServerBox.CheckedItems
    return $allSelectedServers
}

# ==================== EXECUTE BUTTON FUNCTIONS WITH NEW WINDOWS ====================
$btnCoreExecute.Add_Click({
    $selectedServers = $coreServerBox.CheckedItems
    if ($selectedServers.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Please select at least one Core server.", "No Servers Selected", "OK", "Warning")
        return
    }
    
    $selectedTask = $coreTaskDropdown.SelectedItem
	if ($selectedTask -eq "--select command--") {
        [System.Windows.Forms.MessageBox]::Show("Please select a command from the dropdown before executing.", "No Command Selected", "OK", "Warning")
        return
    }
    Log-Message "=== CORE SERVERS TASK EXECUTION STARTED ==="
    Log-Message "Selected Servers: $($selectedServers -join ', ')"
    Log-Message "Selected Task: $selectedTask"
    
    if ($selectedTask -eq "REBOOT Server(s)") {
        Execute-RebootCommand -Servers $selectedServers -ServerType "Core"
        # Reset dropdown back to default after execution
        $coreTaskDropdown.SelectedIndex = 0
        return
    }
    
    $delay = 0
    foreach ($server in $selectedServers) {
        switch ($selectedTask) {
            "STOP APP" {
                $command = "Invoke-Command -ComputerName $server -ScriptBlock {
                    `$imaginetRoot = [System.Environment]::GetEnvironmentVariable('imaginet_root')
                    `$scriptPath = Join-Path `$imaginetRoot 'mst\admin\system5_shutapp.bat'
                    if (Test-Path `$scriptPath) {
                        & `$scriptPath '-F'
                        return 'system5_shutapp.bat -F executed successfully on ' + `$env:COMPUTERNAME
                    } else {
                        return 'system5_shutapp.bat not found at ' + `$scriptPath + ' on ' + `$env:COMPUTERNAME
                    }
                }"
                Execute-CommandInNewWindow -Server $server -Command $command -WindowTitle "CORE $server - STOP APP" -Delay $delay
            }
            "START APP" {
                $command = "Invoke-Command -ComputerName $server -ScriptBlock {
                    `$imaginetRoot = [System.Environment]::GetEnvironmentVariable('imaginet_root')
                    `$scriptPath = Join-Path `$imaginetRoot 'mst\admin\system5_startapp.bat'
                    if (Test-Path `$scriptPath) {
                        & `$scriptPath
                        return 'system5_startapp.bat executed successfully on ' + `$env:COMPUTERNAME
                    } else {
                        return 'system5_startapp.bat not found at ' + `$scriptPath + ' on ' + `$env:COMPUTERNAME
                    }
                }"
                Execute-CommandInNewWindow -Server $server -Command $command -WindowTitle "CORE $server - START APP" -Delay $delay
            }
            "STOP APP/DB" {
                $command = "Invoke-Command -ComputerName $server -ScriptBlock {
                    `$imaginetRoot = [System.Environment]::GetEnvironmentVariable('imaginet_root')
                    `$scriptPath = Join-Path `$imaginetRoot 'mst\admin\system5_shutdown.bat'
                    if (Test-Path `$scriptPath) {
                        & `$scriptPath
                        return 'system5_shutdown.bat executed successfully on ' + `$env:COMPUTERNAME
                    } else {
                        return 'system5_shutdown.bat not found at ' + `$scriptPath + ' on ' + `$env:COMPUTERNAME
                    }
                }"
                Execute-CommandInNewWindow -Server $server -Command $command -WindowTitle "CORE $server - STOP APP/DB" -Delay $delay
            }
            "START APP/DB" {
                $command = "Invoke-Command -ComputerName $server -ScriptBlock {
                    `$imaginetRoot = [System.Environment]::GetEnvironmentVariable('imaginet_root')
                    `$scriptPath = Join-Path `$imaginetRoot 'mst\admin\system5_startup.bat'
                    if (Test-Path `$scriptPath) {
                        & `$scriptPath
                        return 'system5_startup.bat executed successfully on ' + `$env:COMPUTERNAME
                    } else {
                        return 'system5_startup.bat not found at ' + `$scriptPath + ' on ' + `$env:COMPUTERNAME
                    }
                }"
                Execute-CommandInNewWindow -Server $server -Command $command -WindowTitle "CORE $server - START APP/DB" -Delay $delay
            }
        }
        $delay += 2  # 2 second delay between each server
    }
    Log-Message "=== CORE SERVERS TASK EXECUTION COMPLETED ==="
    # Reset dropdown back to default after execution
    $coreTaskDropdown.SelectedIndex = 0
})

#==================== DA SERVERS TASK EXECUTION ====================
$btnDAExecute.Add_Click({
    $selectedServers = $daServerBox.CheckedItems
    if ($selectedServers.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Please select at least one DA server.", "No Servers Selected", "OK", "Warning")
        return
    }
    
    # Get selected task
    $selectedTask = $daTaskDropdown.SelectedItem
	if ($selectedTask -eq "--select command--") {
        [System.Windows.Forms.MessageBox]::Show("Please select a command from the dropdown before executing.", "No Command Selected", "OK", "Warning")
        return
    }
    Log-Message "=== DA SERVERS TASK EXECUTION STARTED ==="
    Log-Message "Selected Servers: $($selectedServers -join ', ')"
    Log-Message "Selected Task: $selectedTask"
    
    if ($selectedTask -eq "REBOOT Server(s)") {
        Execute-RebootCommand -Servers $selectedServers -ServerType "DA"
        # Reset dropdown back to default after execution
        $daTaskDropdown.SelectedIndex = 0
        return
    }
    
    # Execute selected task on each server with staggered delay
    $delay = 0
    foreach ($server in $selectedServers) {
        switch ($selectedTask) {
            "STOP APP" {
                $command = "Invoke-Command -ComputerName $server -ScriptBlock {
                    `$imaginetRoot = [System.Environment]::GetEnvironmentVariable('imaginet_root')
                    `$scriptPath = Join-Path `$imaginetRoot 'mst\admin\system5_shutapp.bat'
                    if (Test-Path `$scriptPath) {
                        & `$scriptPath '-F'
                        return 'system5_shutapp.bat -F executed successfully on ' + `$env:COMPUTERNAME
                    } else {
                        return 'system5_shutapp.bat not found at ' + `$scriptPath + ' on ' + `$env:COMPUTERNAME
                    }
                }"
                Execute-CommandInNewWindow -Server $server -Command $command -WindowTitle "DA $server - STOP APP" -Delay $delay
            }
            "START APP" {
                $command = "Invoke-Command -ComputerName $server -ScriptBlock {
                    `$imaginetRoot = [System.Environment]::GetEnvironmentVariable('imaginet_root')
                    `$scriptPath = Join-Path `$imaginetRoot 'mst\admin\system5_startapp.bat'
                    if (Test-Path `$scriptPath) {
                        & `$scriptPath
                        return 'system5_startapp.bat executed successfully on ' + `$env:COMPUTERNAME
                    } else {
                        return 'system5_startapp.bat not found at ' + `$scriptPath + ' on ' + `$env:COMPUTERNAME
                    }
                }"
                Execute-CommandInNewWindow -Server $server -Command $command -WindowTitle "DA $server - START APP" -Delay $delay
            }
            "STOP APP/DB" {
                $command = "Invoke-Command -ComputerName $server -ScriptBlock {
                    `$imaginetRoot = [System.Environment]::GetEnvironmentVariable('imaginet_root')
                    `$scriptPath = Join-Path `$imaginetRoot 'mst\admin\system5_shutdown.bat'
                    if (Test-Path `$scriptPath) {
                        & `$scriptPath
                        return 'system5_shutdown.bat executed successfully on ' + `$env:COMPUTERNAME
                    } else {
                        return 'system5_shutdown.bat not found at ' + `$scriptPath + ' on ' + `$env:COMPUTERNAME
                    }
                }"
                Execute-CommandInNewWindow -Server $server -Command $command -WindowTitle "DA $server - STOP APP/DB" -Delay $delay
            }
            "START APP/DB" {
                $command = "Invoke-Command -ComputerName $server -ScriptBlock {
                    `$imaginetRoot = [System.Environment]::GetEnvironmentVariable('imaginet_root')
                    `$scriptPath = Join-Path `$imaginetRoot 'mst\admin\system5_startup.bat'
                    if (Test-Path `$scriptPath) {
                        & `$scriptPath
                        return 'system5_startup.bat executed successfully on ' + `$env:COMPUTERNAME
                    } else {
                        return 'system5_startup.bat not found at ' + `$scriptPath + ' on ' + `$env:COMPUTERNAME
                    }
                }"
                Execute-CommandInNewWindow -Server $server -Command $command -WindowTitle "DA $server - START APP/DB" -Delay $delay
            }
        }
        $delay += 2  # 2 second delay between each server
    }
    Log-Message "=== DA SERVERS TASK EXECUTION COMPLETED ==="
    # Reset dropdown back to default after execution
    $daTaskDropdown.SelectedIndex = 0
})

#==================== PORTAL SERVERS TASK EXECUTION ====================
$btnPortalExecute.Add_Click({
    $selectedServers = $portalServerBox.CheckedItems
    if ($selectedServers.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Please select at least one Portal server.", "No Servers Selected", "OK", "Warning")
        return
    }
    
    # Get selected task
    $selectedTask = $portalTaskDropdown.SelectedItem
	if ($selectedTask -eq "--select command--") {
        [System.Windows.Forms.MessageBox]::Show("Please select a command from the dropdown before executing.", "No Command Selected", "OK", "Warning")
        return
    }
    Log-Message "=== PORTAL SERVERS TASK EXECUTION STARTED ==="
    Log-Message "Selected Servers: $($selectedServers -join ', ')"
    Log-Message "Selected Task: $selectedTask"
    
    if ($selectedTask -eq "REBOOT Server(s)") {
        Execute-RebootCommand -Servers $selectedServers -ServerType "Portal"
        # Reset dropdown back to default after execution
        $portalTaskDropdown.SelectedIndex = 0
        return
    }
    # Execute selected task on each server with staggered delay
    $delay = 0
    foreach ($server in $selectedServers) {
        switch ($selectedTask) {
            "Portal STOP" {
                $command = "Invoke-Command -ComputerName $server -ScriptBlock {
                    `$imaginetRoot = [System.Environment]::GetEnvironmentVariable('imaginet_root')
                    `$scriptPath = Join-Path `$imaginetRoot 'Portal\scripts\portal_stop.bat'
                    if (Test-Path `$scriptPath) {
                        & `$scriptPath
                        return 'portal_stop.bat executed successfully on ' + `$env:COMPUTERNAME
                    } else {
                        return 'portal_stop.bat not found at ' + `$scriptPath + ' on ' + `$env:COMPUTERNAME
                    }
                }"
                Execute-CommandInNewWindow -Server $server -Command $command -WindowTitle "PORTAL $server - Portal STOP" -Delay $delay
            }
            "Portal START" {
                $command = "Invoke-Command -ComputerName $server -ScriptBlock {
                    `$imaginetRoot = [System.Environment]::GetEnvironmentVariable('imaginet_root')
                    `$scriptPath = Join-Path `$imaginetRoot 'Portal\scripts\portal_start.bat'
                    if (Test-Path `$scriptPath) {
                        & `$scriptPath
                        return 'portal_start.bat executed successfully on ' + `$env:COMPUTERNAME
                    } else {
                        return 'portal_start.bat not found at ' + `$scriptPath + ' on ' + `$env:COMPUTERNAME
                    }
                }"
                Execute-CommandInNewWindow -Server $server -Command $command -WindowTitle "PORTAL $server - Portal START" -Delay $delay
            }
            "Portal RESTART" {
                $command = "Invoke-Command -ComputerName $server -ScriptBlock {
                    `$imaginetRoot = [System.Environment]::GetEnvironmentVariable('imaginet_root')
                    `$scriptPath = Join-Path `$imaginetRoot 'Portal\scripts\portal_restart.bat'
                    if (Test-Path `$scriptPath) {
                        & `$scriptPath
                        return 'portal_restart.bat executed successfully on ' + `$env:COMPUTERNAME
                    } else {
                        return 'portal_restart.bat not found at ' + `$scriptPath + ' on ' + `$env:COMPUTERNAME
                    }
                }"
                Execute-CommandInNewWindow -Server $server -Command $command -WindowTitle "PORTAL $server - Portal RESTART" -Delay $delay
            }
        }
        $delay += 2  # 2 second delay between each server
    }
    Log-Message "=== PORTAL SERVERS TASK EXECUTION COMPLETED ==="
    # Reset dropdown back to default after execution
    $portalTaskDropdown.SelectedIndex = 0
})

#==================== OTHER SERVERS TASK EXECUTION ====================
$btnOtherExecute.Add_Click({
    $selectedServers = $otherServerBox.CheckedItems
    if ($selectedServers.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Please select at least one Other server.", "No Servers Selected", "OK", "Warning")
        return
    }
    
    # Get selected task
    $selectedTask = $otherTaskDropdown.SelectedItem
	if ($selectedTask -eq "--select command--") {
        [System.Windows.Forms.MessageBox]::Show("Please select a command from the dropdown before executing.", "No Command Selected", "OK", "Warning")
        return
    }
    Log-Message "=== OTHER SERVERS TASK EXECUTION STARTED ==="
    Log-Message "Selected Servers: $($selectedServers -join ', ')"
    Log-Message "Selected Task: $selectedTask"
    
    if ($selectedTask -eq "REBOOT Server(s)") {
        Execute-RebootCommand -Servers $selectedServers -ServerType "Other"
        # Reset dropdown back to default after execution
        $otherTaskDropdown.SelectedIndex = 0
        return
    }
    
    # Execute selected task on each server with staggered delay
    $delay = 0
    foreach ($server in $selectedServers) {
        switch ($selectedTask) {
            "ES RESTART" {
                $command = "Invoke-Command -ComputerName $server -ScriptBlock {
                    `$services = @('ElasticSearch-service', 'ES-Cerebro-service')
                    `$results = @()
                    
                    # Stop services
                    foreach (`$service in `$services) {
                        try {
                            if (Get-Service -Name `$service -ErrorAction SilentlyContinue) {
                                Stop-Service -Name `$service -Force -ErrorAction Stop
                                `$results += 'Service ' + `$service + ' stopped successfully on ' + `$env:COMPUTERNAME
                            } else {
                                `$results += 'Service ' + `$service + ' not found on ' + `$env:COMPUTERNAME
                            }
                        } catch {
                            `$results += 'ERROR stopping service ' + `$service + ' on ' + `$env:COMPUTERNAME + ' - ' + `$_
                        }
                    }
                    
                    # Wait a moment
                    Start-Sleep -Seconds 3
                    
                    # Start services
                    foreach (`$service in `$services) {
                        try {
                            if (Get-Service -Name `$service -ErrorAction SilentlyContinue) {
                                Start-Service -Name `$service -ErrorAction Stop
                                `$results += 'Service ' + `$service + ' started successfully on ' + `$env:COMPUTERNAME
                            } else {
                                `$results += 'Service ' + `$service + ' not found on ' + `$env:COMPUTERNAME
                            }
                        } catch {
                            `$results += 'ERROR starting service ' + `$service + ' on ' + `$env:COMPUTERNAME + ' - ' + `$_
                        }
                    }
                    
                    return `$results -join '`n'
                }"
                Execute-CommandInNewWindow -Server $server -Command $command -WindowTitle "OTHER $server - ES RESTART" -Delay $delay
            }
            "EA STOP" {
                $command = "Invoke-Command -ComputerName $server -ScriptBlock {
                    `$imaginetRoot = [System.Environment]::GetEnvironmentVariable('imaginet_root')
                    `$scriptPath = Join-Path `$imaginetRoot \Analytics\Scripts\'event_utils.py'
                    if (Test-Path `$scriptPath) {
                        & python `$scriptPath '-action' 'stop'
                        return 'event_utils.py -action stop executed successfully on ' + `$env:COMPUTERNAME
                    } else {
                        return 'event_utils.py not found at ' + `$scriptPath + ' on ' + `$env:COMPUTERNAME
                    }
                }"
                Execute-CommandInNewWindow -Server $server -Command $command -WindowTitle "OTHER $server - EA STOP" -Delay $delay
            }
            "EA START" {
                $command = "Invoke-Command -ComputerName $server -ScriptBlock {
                    `$imaginetRoot = [System.Environment]::GetEnvironmentVariable('imaginet_root')
                    `$scriptPath = Join-Path `$imaginetRoot \Analytics\Scripts\'event_utils.py'
                    if (Test-Path `$scriptPath) {
                        & python `$scriptPath '-action' 'start'
                        return 'event_utils.py -action start executed successfully on ' + `$env:COMPUTERNAME
                    } else {
                        return 'event_utils.py not found at ' + `$scriptPath + ' on ' + `$env:COMPUTERNAME
                    }
                }"
                Execute-CommandInNewWindow -Server $server -Command $command -WindowTitle "OTHER $server - EA START" -Delay $delay
            }
            "EA RESTART" {
                $command = "Invoke-Command -ComputerName $server -ScriptBlock {
                    `$imaginetRoot = [System.Environment]::GetEnvironmentVariable('imaginet_root')
                    `$scriptPath = Join-Path `$imaginetRoot \Analytics\Scripts\'event_utils.py'
                    if (Test-Path `$scriptPath) {
                        & python `$scriptPath '-action' 'restart'
                        return 'event_utils.py -action restart executed successfully on ' + `$env:COMPUTERNAME
                    } else {
                        return 'event_utils.py not found at ' + `$scriptPath + ' on ' + `$env:COMPUTERNAME
                    }
                }"
                Execute-CommandInNewWindow -Server $server -Command $command -WindowTitle "OTHER $server - EA RESTART" -Delay $delay
            }
            "EA STATUS" {
                $command = "Invoke-Command -ComputerName $server -ScriptBlock {
                    `$imaginetRoot = [System.Environment]::GetEnvironmentVariable('imaginet_root')
                    `$scriptPath = Join-Path `$imaginetRoot \Analytics\Scripts\'event_utils.py'
                    if (Test-Path `$scriptPath) {
                        & python `$scriptPath '-action' 'status'
                        return 'event_utils.py -action status executed successfully on ' + `$env:COMPUTERNAME
                    } else {
                        return 'event_utils.py not found at ' + `$scriptPath + ' on ' + `$env:COMPUTERNAME
                    }
                }"
                Execute-CommandInNewWindow -Server $server -Command $command -WindowTitle "OTHER $server - EA STATUS" -Delay $delay
            }
            "RA STOP" {
                $command = "Invoke-Command -ComputerName $server -ScriptBlock {
                    `$imaginetRoot = [System.Environment]::GetEnvironmentVariable('imaginet_root')
                    `$scriptPath = Join-Path `$imaginetRoot 'scripts\ReportAnalytics.bat'
                    if (Test-Path `$scriptPath) {
                        & `$scriptPath 'stop'
                        return 'ReportAnalytics.bat stop executed successfully on ' + `$env:COMPUTERNAME
                    } else {
                        return 'ReportAnalytics.bat not found at ' + `$scriptPath + ' on ' + `$env:COMPUTERNAME
                    }
                }"
                Execute-CommandInNewWindow -Server $server -Command $command -WindowTitle "OTHER $server - RA STOP" -Delay $delay
            }
            "RA START" {
                $command = "Invoke-Command -ComputerName $server -ScriptBlock {
                    `$imaginetRoot = [System.Environment]::GetEnvironmentVariable('imaginet_root')
                    `$scriptPath = Join-Path `$imaginetRoot 'scripts\ReportAnalytics.bat'
                    if (Test-Path `$scriptPath) {
                        & `$scriptPath 'start'
                        return 'ReportAnalytics.bat start executed successfully on ' + `$env:COMPUTERNAME
                    } else {
                        return 'ReportAnalytics.bat not found at ' + `$scriptPath + ' on ' + `$env:COMPUTERNAME
                    }
                }"
                Execute-CommandInNewWindow -Server $server -Command $command -WindowTitle "OTHER $server - RA START" -Delay $delay
            }
            "RA RESTART" {
                $command = "Invoke-Command -ComputerName $server -ScriptBlock {
                    `$imaginetRoot = [System.Environment]::GetEnvironmentVariable('imaginet_root')
                    `$scriptPath = Join-Path `$imaginetRoot 'scripts\ReportAnalytics.bat'
                    if (Test-Path `$scriptPath) {
                        & `$scriptPath 'restart'
                        return 'ReportAnalytics.bat restart executed successfully on ' + `$env:COMPUTERNAME
                    } else {
                        return 'ReportAnalytics.bat not found at ' + `$scriptPath + ' on ' + `$env:COMPUTERNAME
                    }
                }"
                Execute-CommandInNewWindow -Server $server -Command $command -WindowTitle "OTHER $server - RA RESTART" -Delay $delay
            }
        }
        $delay += 2  # 2 second delay between each server
    }
    Log-Message "=== OTHER SERVERS TASK EXECUTION COMPLETED ==="
    # Reset dropdown back to default after execution
    $otherTaskDropdown.SelectedIndex = 0
})

# ==================== SERVICES RESTART SECTION ====================
$btnServiceRestart.Add_Click({
    if (-not (Validate-Servers)) { return }
    
    $selectedServers = Get-AllSelectedServers
    $selectedServices = $serviceBox.CheckedItems
    
    if ($selectedServices.Count -eq 0) {
        Log-Message "WARNING: No services selected for restart."
        return
    }
    
    Log-Message "=== SERVICES RESTART OPERATION STARTED ==="
    Log-Message "Selected Servers: $($selectedServers -join ', ')"
    Log-Message "Selected Services: $($selectedServices -join ', ')"
    
    foreach ($Server in $selectedServers) {
        Log-Message "Processing Services on ${Server}..."
        
        foreach ($Service in $selectedServices) {
            try {
                Log-Message "Running: Restart-Service -Name $Service on $Server"
                $result = Invoke-Command -ComputerName $Server -ScriptBlock {
                    Param ($Service)
                    if (Get-Service -Name $Service -ErrorAction SilentlyContinue) {
                        Restart-Service -Name $Service -Force -ErrorAction Stop
                        return "${Service} restarted successfully on $env:COMPUTERNAME."
                    } else {
                        return "Service ${Service} not found on $env:COMPUTERNAME."
                    }
                } -ArgumentList $Service -ErrorAction Stop
                Log-Message $result
            } catch {
                Log-Message "ERROR restarting service ${Service} on ${Server} - $_"
            }
        }
        Log-Message "Services processing completed for ${Server}"
    }
    Log-Message "=== SERVICES RESTART OPERATION COMPLETED ==="
})

# ==================== SERVICES STOP SECTION ====================
$btnServiceStop.Add_Click({
    if (-not (Validate-Servers)) { return }
    
    $selectedServers = Get-AllSelectedServers
    $selectedServices = $serviceBox.CheckedItems
    
    if ($selectedServices.Count -eq 0) {
        Log-Message "WARNING: No services selected for stop operation."
        return
    }
    
    Log-Message "=== SERVICES STOP OPERATION STARTED ==="
    Log-Message "Selected Servers: $($selectedServers -join ', ')"
    Log-Message "Selected Services: $($selectedServices -join ', ')"
    
    foreach ($Server in $selectedServers) {
        Log-Message "Stopping Services on ${Server}..."
        
        foreach ($Service in $selectedServices) {
            try {
                Log-Message "Running: Stop-Service -Name $Service on $Server"
                $result = Invoke-Command -ComputerName $Server -ScriptBlock {
                    Param ($Service)
                    if (Get-Service -Name $Service -ErrorAction SilentlyContinue) {
                        Stop-Service -Name $Service -Force -ErrorAction Stop
                        return "${Service} stopped successfully on $env:COMPUTERNAME."
                    } else {
                        return "Service ${Service} not found on $env:COMPUTERNAME."
                    }
                } -ArgumentList $Service -ErrorAction Stop
                Log-Message $result
            } catch {
                Log-Message "ERROR stopping service ${Service} on ${Server} - $_"
            }
        }
        Log-Message "Services stop processing completed for ${Server}"
    }
    Log-Message "=== SERVICES STOP OPERATION COMPLETED ==="
})

# ==================== SERVICES START SECTION ====================
$btnServiceStart.Add_Click({
    if (-not (Validate-Servers)) { return }
    
    $selectedServers = Get-AllSelectedServers
    $selectedServices = $serviceBox.CheckedItems
    
    if ($selectedServices.Count -eq 0) {
        Log-Message "WARNING: No services selected for start operation."
        return
    }
    
    Log-Message "=== SERVICES START OPERATION STARTED ==="
    Log-Message "Selected Servers: $($selectedServers -join ', ')"
    Log-Message "Selected Services: $($selectedServices -join ', ')"
    
    foreach ($Server in $selectedServers) {
        Log-Message "Starting Services on ${Server}..."
        
        foreach ($Service in $selectedServices) {
            try {
                Log-Message "Running: Start-Service -Name $Service on $Server"
                $result = Invoke-Command -ComputerName $Server -ScriptBlock {
                    Param ($Service)
                    if (Get-Service -Name $Service -ErrorAction SilentlyContinue) {
                        Start-Service -Name $Service -ErrorAction Stop
                        return "${Service} started successfully on $env:COMPUTERNAME."
                    } else {
                        return "Service ${Service} not found on $env:COMPUTERNAME."
                    }
                } -ArgumentList $Service -ErrorAction Stop
                Log-Message $result
            } catch {
                Log-Message "ERROR starting service ${Service} on ${Server} - $_"
            }
        }
        Log-Message "Services start processing completed for ${Server}"
    }
    Log-Message "=== SERVICES START OPERATION COMPLETED ==="
})

# ==================== SERVICES TASKKILL SECTION ====================
$btnServiceTaskKill.Add_Click({
    if (-not (Validate-Servers)) { return }
    
    $selectedServers = Get-AllSelectedServers
    $selectedServices = $serviceBox.CheckedItems
    
    if ($selectedServices.Count -eq 0) {
        Log-Message "WARNING: No services selected for taskkill operation."
        return
    }
    
    $confirmResult = [System.Windows.Forms.MessageBox]::Show(
        "Are you sure you want to force kill the selected services? This action cannot be undone.", 
        "Confirm TASKKILL Operation", 
        "YesNo", 
        "Warning"
    )
    
    if ($confirmResult -eq "No") {
        Log-Message "TASKKILL operation cancelled by user."
        return
    }
    
    Log-Message "=== SERVICES TASKKILL OPERATION STARTED ==="
    Log-Message "Selected Servers: $($selectedServers -join ', ')"
    Log-Message "Selected Services: $($selectedServices -join ', ')"
    
    foreach ($Server in $selectedServers) {
        Log-Message "Force killing Services on ${Server}..."
        
        foreach ($Service in $selectedServices) {
            try {
                Log-Message "Running: TASKKILL /F /IM $Service on $Server"
                $result = Invoke-Command -ComputerName $Server -ScriptBlock {
                    Param ($Service)
                    $output = & taskkill /F /IM "$Service.exe" 2>&1
                        return "TASKKILL output for $($Service): $output"
                } -ArgumentList $Service -ErrorAction Stop
                Log-Message $result
            } catch {
                Log-Message "ERROR with taskkill for ${Service} on ${Server} - $_"
            }
        }
        Log-Message "TASKKILL processing completed for ${Server}"
    }
    Log-Message "=== SERVICES TASKKILL OPERATION COMPLETED ==="
})

# ==================== SERVICES GET INFO SECTION ====================
$btnServiceGetInfo.Add_Click({
    if (-not (Validate-Servers)) { return }
    
    $selectedServers = Get-AllSelectedServers
    $selectedServices = $serviceBox.CheckedItems
    
    if ($selectedServices.Count -eq 0) {
        Log-Message "WARNING: No services selected for GET INFO operation."
        return
    }
    
    Log-Message "=== SERVICES GET INFO OPERATION STARTED ==="
    Log-Message "Selected Servers: $($selectedServers -join ', ')"
    Log-Message "Selected Services: $($selectedServices -join ', ')"
    
    # Show progress message
    $progressForm = New-Object System.Windows.Forms.Form
    $progressForm.Text = "Gathering Service Information..."
    $progressForm.Size = New-Object System.Drawing.Size(400, 150)
    $progressForm.StartPosition = "CenterParent"
    $progressForm.FormBorderStyle = "FixedDialog"
    $progressForm.MaximizeBox = $false
    $progressForm.MinimizeBox = $false
    
    $progressLabel = New-Object System.Windows.Forms.Label
    $progressLabel.Text = "Please wait while gathering service information from servers..."
    $progressLabel.Location = New-Object System.Drawing.Point(20, 30)
    $progressLabel.Size = New-Object System.Drawing.Size(350, 60)
    $progressLabel.TextAlign = "MiddleCenter"
    $progressForm.Controls.Add($progressLabel)
    
    $progressForm.Show()
    $progressForm.Refresh()
    
    # Build formatted report
    $report = @()
    $report += "=== SERVICES INFORMATION REPORT ==="
    $report += "Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    $report += "Servers: $($selectedServers.Count) selected"
    $report += "Services: $($selectedServices.Count) selected - $($selectedServices -join ', ')"
    $report += ""
    
    foreach ($Server in $selectedServers | Sort-Object) {
        try {
            $progressLabel.Text = "Gathering service info from: $Server"
            $progressForm.Refresh()
            
            Log-Message "Getting service info from $Server..."
            
            $result = Invoke-Command -ComputerName $Server -ScriptBlock {
                Param ($SelectedServices)
                
                $results = @()
                $results += "=========================================="
                $results += "SERVER: $env:COMPUTERNAME"
                $results += "=========================================="
                $results += ""
                
                foreach ($serviceName in $SelectedServices) {
                    $results += "--- SERVICE: $serviceName ---"
                    
                    # Get service information
                    try {
                        $service = Get-Service -Name $serviceName -ErrorAction Stop
                        $results += "Display Name: $($service.DisplayName)"
                        $results += "Status: $($service.Status)"
                        $results += "Start Type: $($service.StartType)"
                        
                        # Get process information if service is running
                        if ($service.Status -eq "Running") {
                            try {
                                # Get the service process ID
                                $serviceWMI = Get-WmiObject Win32_Service -Filter "Name='$serviceName'" -ErrorAction Stop
                                if ($serviceWMI.ProcessId -and $serviceWMI.ProcessId -gt 0) {
                                    $results += "Process ID (PID): **$($serviceWMI.ProcessId)**"
                                    
                                    # Get process details
                                    $process = Get-Process -Id $serviceWMI.ProcessId -ErrorAction Stop
                                    $memoryMB = [math]::Round($process.WorkingSet / 1MB, 2)
                                    $results += "Process Name: $($process.ProcessName)"
                                    $results += "Memory Usage: $memoryMB MB"
                                    $results += "CPU Time: $($process.TotalProcessorTime.ToString('hh\:mm\:ss'))"
                                    $results += "Start Time: $($process.StartTime)"
                                    $results += "Thread Count: $($process.Threads.Count)"
                                    $results += "Handle Count: $($process.HandleCount)"
                                    
                                    # Get executable path
                                    try {
                                        $procWMI = Get-WmiObject Win32_Process -Filter "ProcessId=$($serviceWMI.ProcessId)" -ErrorAction Stop
                                        if ($procWMI.ExecutablePath) {
                                            $results += "Executable Path: $($procWMI.ExecutablePath)"
                                        }
                                    } catch {
                                        $results += "Executable Path: Unable to retrieve"
                                    }
                                } else {
                                    $results += "Process ID: Not available (shared process or service host)"
                                    
                                    # Try to find if it's in a service host
                                    try {
                                        $svchost = Get-WmiObject Win32_Service -Filter "Name='$serviceName'" | Where-Object { $_.PathName -like "*svchost*" }
                                        if ($svchost) {
                                            $results += "Service Type: Hosted in svchost.exe"
                                            if ($svchost.ProcessId -gt 0) {
                                                $results += "Host Process ID: **$($svchost.ProcessId)**"
                                            }
                                        }
                                    } catch {
                                        # Ignore if can't get svchost info
                                    }
                                }
                            } catch {
                                $results += "Process Info: Error retrieving process details - $_"
                            }
                        } else {
                            $results += "Process ID: Service not running"
                        }
                        
                        # Get service dependencies
                        if ($service.ServicesDependedOn.Count -gt 0) {
                            $results += "Depends On: $($service.ServicesDependedOn.Name -join ', ')"
                        }
                        if ($service.DependentServices.Count -gt 0) {
                            $results += "Dependents: $($service.DependentServices.Name -join ', ')"
                        }
                        
                        # Get service account
                        try {
                            $serviceWMI2 = Get-WmiObject Win32_Service -Filter "Name='$serviceName'" -ErrorAction Stop
                            if ($serviceWMI2.StartName) {
                                $results += "Log On As: $($serviceWMI2.StartName)"
                            }
                        } catch {
                            $results += "Log On As: Unable to retrieve"
                        }
                        
                    } catch {
                        $results += "ERROR: Service '$serviceName' not found or access denied - $_"
                    }
                    
                    $results += ""
                }
                
                return @{
                    Success = $true
                    Data = $results
                    Hostname = $env:COMPUTERNAME
                }
            } -ArgumentList (,$selectedServices) -ErrorAction Stop
            
            if ($result.Success) {
                $report += $result.Data
                $report += ""
                Log-Message "Successfully retrieved service info from $Server"
            } else {
                $report += "ERROR getting service info from $Server"
                $report += ""
                Log-Message "ERROR getting service info from $Server"
            }
            
        } catch {
            $report += "ERROR connecting to $Server - $_"
            $report += ""
            Log-Message "ERROR connecting to $Server - $_"
        }
    }
    
    $progressForm.Close()
    
    # Create results window
    $resultsForm = New-Object System.Windows.Forms.Form
    $resultsForm.Text = "Service Information Report - $($selectedServers.Count) Servers, $($selectedServices.Count) Services"
    $resultsForm.Size = New-Object System.Drawing.Size(1250, 700)
    $resultsForm.StartPosition = "CenterParent"
    $resultsForm.FormBorderStyle = "Sizable"
    
	# Text box for results
	$resultsTextBox = New-Object System.Windows.Forms.TextBox
	$resultsTextBox.Location = New-Object System.Drawing.Point(10, 10)
	$resultsTextBox.Size = New-Object System.Drawing.Size(1210, 600)
	$resultsTextBox.Multiline = $true
	$resultsTextBox.ScrollBars = "Both"
	$resultsTextBox.Font = New-Object System.Drawing.Font("Consolas", 9)
	$resultsTextBox.WordWrap = $false
	$resultsTextBox.ReadOnly = $true
	$resultsTextBox.Text = ($report -join "`r`n") -replace '\*\*(\d+)\*\*', '$1'
	$resultsTextBox.SelectionStart = 0
	$resultsTextBox.SelectionLength = 0
	$resultsTextBox.Anchor = "Top, Left, Bottom, Right"
	$resultsForm.Controls.Add($resultsTextBox)	

	# Copy to Clipboard button
    $copyButton = New-Object System.Windows.Forms.Button
    $copyButton.Text = "Copy to Clipboard"
    $copyButton.Location = New-Object System.Drawing.Point(350, 620)
    $copyButton.Size = New-Object System.Drawing.Size(120, 30)
    $copyButton.Anchor = "Bottom"
    $copyButton.BackColor = [System.Drawing.Color]::LightGreen
    $copyButton.Font = New-Object System.Drawing.Font("Arial", 9, [System.Drawing.FontStyle]::Bold)
    $copyButton.Add_Click({
        try {
            [System.Windows.Forms.Clipboard]::SetText($resultsTextBox.Text)
            $copyButton.Text = "Copied!"
            $copyButton.BackColor = [System.Drawing.Color]::Gold
            $timer = New-Object System.Windows.Forms.Timer
            $timer.Interval = 1500
            $timer.Add_Tick({
                $copyButton.Text = "Copy to Clipboard"
                $copyButton.BackColor = [System.Drawing.Color]::LightGreen
                $timer.Stop()
                $timer.Dispose()
            })
            $timer.Start()
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Failed to copy to clipboard: $_", "Copy Error", "OK", "Error")
        }
    })
    $resultsForm.Controls.Add($copyButton)
    
    # Close button
    $closeButton = New-Object System.Windows.Forms.Button
    $closeButton.Text = "Close"
    $closeButton.Location = New-Object System.Drawing.Point(480, 620)
    $closeButton.Size = New-Object System.Drawing.Size(80, 30)
    $closeButton.Anchor = "Bottom"
    $closeButton.Add_Click({
        $resultsForm.Close()
    })
    $resultsForm.Controls.Add($closeButton)
    
    # Show results window
    $resultsForm.ShowDialog()
    
    Log-Message "=== SERVICES GET INFO OPERATION COMPLETED ==="
})
#
function Invoke-JavaInfo {
    if (-not (Validate-Servers)) { return }
    $selectedServers = Get-AllSelectedServers
    Log-Message "=== JAVA INFO OPERATION STARTED ==="
    Log-Message "Selected Servers: $($selectedServers -join ', ')"

    $progressForm = New-Object System.Windows.Forms.Form
    $progressForm.Text = "Gathering Java Information..."
    $progressForm.Size = New-Object System.Drawing.Size(400, 150)
    $progressForm.StartPosition = "CenterParent"
    $progressForm.FormBorderStyle = "FixedDialog"
    $progressForm.MaximizeBox = $false
    $progressForm.MinimizeBox = $false
    $progressLabel = New-Object System.Windows.Forms.Label
    $progressLabel.Text = "Please wait while gathering Java process information..."
    $progressLabel.Location = New-Object System.Drawing.Point(20, 30)
    $progressLabel.Size = New-Object System.Drawing.Size(350, 60)
    $progressLabel.TextAlign = "MiddleCenter"
    $progressForm.Controls.Add($progressLabel)
    [void]$progressForm.Show(); $progressForm.Refresh()

    $report = @()
    $report += "=== JAVA PROCESS INFORMATION REPORT ==="
    $report += "Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    $report += "Servers: $($selectedServers.Count) selected"
    $report += ""

    foreach ($Server in $selectedServers | Sort-Object) {
        try {
            $progressLabel.Text = "Gathering Java info from: $Server"; $progressForm.Refresh()
            Log-Message "Getting Java process info from $Server..."

            $result = Invoke-Command -ComputerName $Server -ScriptBlock {
                try {
                    $cmd = 'Get-CimInstance Win32_Process | where Name -like *java*'
                    $javaProcesses = Get-CimInstance Win32_Process | Where-Object { $_.Name -like "*java*" } | Select-Object ProcessId, Name, @{Name="Memory(MB)";Expression={[math]::Round($_.WorkingSetSize/1MB,2)}}, @{Name="AppName";Expression={if($_.CommandLine -match '-Dapp\.name=([^\s]+)'){ $matches[1] } else { "N/A" }}} | Sort-Object AppName
                    $results = @()
                    $results += "=========================================="
                    $results += "SERVER: $env:COMPUTERNAME"
                    $results += "=========================================="
                    $results += "Command: $cmd"
                    $results += ""
                    if ($javaProcesses) {
                        $results += "ProcessId".PadRight(12) + "Name".PadRight(25) + "Memory(MB)".PadRight(15) + "AppName"
                        $results += "-" * 80
                        foreach ($proc in $javaProcesses) {
                            $line = $proc.ProcessId.ToString().PadRight(12) + $proc.Name.PadRight(25) + $proc.'Memory(MB)'.ToString().PadRight(15) + $proc.AppName
                            $results += $line
                        }
                        $results += ""
                        $results += "Total Java Processes: $($javaProcesses.Count)"
                        $totalMemory = ($javaProcesses | Measure-Object -Property 'Memory(MB)' -Sum).Sum
                        $results += "Total Memory Usage: $([math]::Round($totalMemory, 2)) MB"
                    } else {
                        $results += "No Java processes found on this server."
                    }
                    return @{ Success = $true; Data = $results; Hostname = $env:COMPUTERNAME }
                } catch { return @{ Success = $false; Error = $_.Exception.Message; Hostname = $env:COMPUTERNAME } }
            } -ErrorAction Stop

            if ($result.Success) { $report += $result.Data; $report += ""; Log-Message "Successfully retrieved Java info from $Server" }
            else { $report += "ERROR getting Java info from $Server - $($result.Error)"; $report += ""; Log-Message "ERROR getting Java info from $Server - $($result.Error)" }
        } catch { $report += "ERROR connecting to $Server - $_"; $report += ""; Log-Message "ERROR connecting to $Server - $_" }
    }

    $progressForm.Close()

    $resultsForm = New-Object System.Windows.Forms.Form
    $resultsForm.Text = "Java Process Information Report - $($selectedServers.Count) Servers"
    $resultsForm.Size = New-Object System.Drawing.Size(1250, 700)
    $resultsForm.StartPosition = "CenterParent"
    $resultsForm.FormBorderStyle = "Sizable"
    $resultsTextBox = New-Object System.Windows.Forms.TextBox
    $resultsTextBox.Location = New-Object System.Drawing.Point(10, 10)
    $resultsTextBox.Size = New-Object System.Drawing.Size(1210, 600)
    $resultsTextBox.Multiline = $true
    $resultsTextBox.ScrollBars = "Both"
    $resultsTextBox.Font = New-Object System.Drawing.Font("Consolas", 9)
    $resultsTextBox.WordWrap = $false
    $resultsTextBox.ReadOnly = $true
    $resultsTextBox.Text = $report -join "`r`n"
    $resultsTextBox.SelectionStart = 0; $resultsTextBox.SelectionLength = 0
    $resultsTextBox.Anchor = "Top, Left, Bottom, Right"
    $resultsForm.Controls.Add($resultsTextBox)
    $copyButton = New-Object System.Windows.Forms.Button
    $copyButton.Text = "Copy to Clipboard"
    $copyButton.Location = New-Object System.Drawing.Point(350, 620)
    $copyButton.Size = New-Object System.Drawing.Size(120, 30)
    $copyButton.Anchor = "Bottom"
    $copyButton.BackColor = [System.Drawing.Color]::LightGreen
    $copyButton.Font = New-Object System.Drawing.Font("Arial", 9, [System.Drawing.FontStyle]::Bold)
    $copyButton.Add_Click({ try { [System.Windows.Forms.Clipboard]::SetText($resultsTextBox.Text); $copyButton.Text = "Copied!"; $copyButton.BackColor = [System.Drawing.Color]::Gold; $timer = New-Object System.Windows.Forms.Timer; $timer.Interval = 1500; $timer.Add_Tick({ $copyButton.Text = "Copy to Clipboard"; $copyButton.BackColor = [System.Drawing.Color]::LightGreen; $timer.Stop(); $timer.Dispose() }); $timer.Start() } catch { [System.Windows.Forms.MessageBox]::Show("Failed to copy to clipboard: $_", "Copy Error", "OK", "Error") } })
    $resultsForm.Controls.Add($copyButton)
    $closeButton = New-Object System.Windows.Forms.Button
    $closeButton.Text = "Close"
    $closeButton.Location = New-Object System.Drawing.Point(480, 620)
    $closeButton.Size = New-Object System.Drawing.Size(80, 30)
    $closeButton.Anchor = "Bottom"
    $closeButton.Add_Click({ $resultsForm.Close() })
    $resultsForm.Controls.Add($closeButton)
    [void]$resultsForm.ShowDialog()
    Log-Message "=== JAVA INFO OPERATION COMPLETED ==="
}

# ==================== SERVICES ENABLE SECTION ====================
$btnServiceEnable.Add_Click({
    if (-not (Validate-Servers)) { return }
    
    $selectedServers = Get-AllSelectedServers
    $selectedServices = $serviceBox.CheckedItems
    
    if ($selectedServices.Count -eq 0) {
        Log-Message "WARNING: No services selected for enable operation."
        return
    }
    
    Log-Message "=== SERVICES ENABLE OPERATION STARTED ==="
    Log-Message "Selected Servers: $($selectedServers -join ', ')"
    Log-Message "Selected Services: $($selectedServices -join ', ')"
    
    foreach ($Server in $selectedServers) {
        Log-Message "Enabling Services on ${Server}..."
        
        foreach ($Service in $selectedServices) {
            try {
                Log-Message "Running: Set-Service -Name $Service -StartupType Automatic on $Server"
                $result = Invoke-Command -ComputerName $Server -ScriptBlock {
                    Param ($Service)
                    if (Get-Service -Name $Service -ErrorAction SilentlyContinue) {
                        Set-Service -Name $Service -StartupType Automatic -ErrorAction Stop
                        return "${Service} enabled (set to Automatic startup) successfully on $env:COMPUTERNAME."
                    } else {
                        return "Service ${Service} not found on $env:COMPUTERNAME."
                    }
                } -ArgumentList $Service -ErrorAction Stop
                Log-Message $result
            } catch {
                Log-Message "ERROR enabling service ${Service} on ${Server} - $_"
            }
        }
        Log-Message "Services enable processing completed for ${Server}"
    }
    Log-Message "=== SERVICES ENABLE OPERATION COMPLETED ==="
})

# ==================== SERVICES DISABLE SECTION ====================
$btnServiceDisable.Add_Click({
    if (-not (Validate-Servers)) { return }
    
    $selectedServers = Get-AllSelectedServers
    $selectedServices = $serviceBox.CheckedItems
    
    if ($selectedServices.Count -eq 0) {
        Log-Message "WARNING: No services selected for disable operation."
        return
    }
    
    $confirmResult = [System.Windows.Forms.MessageBox]::Show(
        "Are you sure you want to STOP and DISABLE the selected services? This will prevent them from starting automatically.", 
        "Confirm DISABLE Operation", 
        "YesNo", 
        "Warning"
    )
    
    if ($confirmResult -eq "No") {
        Log-Message "DISABLE operation cancelled by user."
        return
    }
    
    Log-Message "=== SERVICES DISABLE OPERATION STARTED ==="
    Log-Message "Selected Servers: $($selectedServers -join ', ')"
    Log-Message "Selected Services: $($selectedServices -join ', ')"
    
    foreach ($Server in $selectedServers) {
        Log-Message "Disabling Services on ${Server}..."
        
        foreach ($Service in $selectedServices) {
            try {
                Log-Message "Running: Stop-Service and Set-Service -StartupType Disabled for $Service on $Server"
                $result = Invoke-Command -ComputerName $Server -ScriptBlock {
                    Param ($Service)
                    if (Get-Service -Name $Service -ErrorAction SilentlyContinue) {
                        # First stop the service
                        try {
                            Stop-Service -Name $Service -Force -ErrorAction Stop
                            $stopResult = "${Service} stopped successfully"
                        } catch {
                            $stopResult = "Warning: Could not stop ${Service} - $_"
                        }
                        
                        # Then disable it
                        Set-Service -Name $Service -StartupType Disabled -ErrorAction Stop
                        return "$stopResult and ${Service} disabled successfully on $env:COMPUTERNAME."
                    } else {
                        return "Service ${Service} not found on $env:COMPUTERNAME."
                    }
                } -ArgumentList $Service -ErrorAction Stop
                Log-Message $result
            } catch {
                Log-Message "ERROR disabling service ${Service} on ${Server} - $_"
            }
        }
        Log-Message "Services disable processing completed for ${Server}"
    }
    Log-Message "=== SERVICES DISABLE OPERATION COMPLETED ==="
})
		# ==================== PROCESSES RESTART_SERVICE SECTION ====================
		$btnProcessRestart.Add_Click({
			if (-not (Validate-Servers)) { return }
			
			$selectedServers = Get-AllSelectedServers
			$selectedProcesses = $processBox.CheckedItems
			
			if ($selectedProcesses.Count -eq 0) {
				Log-Message "WARNING: No processes selected for restart."
				return
			}
			
			Log-Message "=== PROCESSES RESTART_SERVICE OPERATION STARTED ==="
			Log-Message "Selected Servers: $($selectedServers -join ', ')"
			Log-Message "Selected Processes: $($selectedProcesses -join ', ')"
			
			$delay = 0
			foreach ($server in $selectedServers) {
				foreach ($process in $selectedProcesses) {
					$command = "Invoke-Command -ComputerName $server -ScriptBlock {
						`$Process = '$process'
						`$rstrtSvc = 'restart_service'
						`$prtCmd = '-p'
						`$prtTls = 'all'
						
						`$imaginetRoot = [System.Environment]::GetEnvironmentVariable('imaginet_root')
						`$adminCmdPaths = if (`$imaginetRoot) {
							@(
								(Join-Path `$imaginetRoot 'utils\AdminCommand.exe'),
								(Join-Path `$imaginetRoot 'Portal\Dll\SVRender\AdminCommand.exe')
							)
						} else {
							@('AdminCommand.exe')
						}
						
						`$commandExecuted = `$false
						`$results = @()

						foreach (`$adminCmd in `$adminCmdPaths) {
							try {
								if (`$adminCmd -eq 'AdminCommand.exe' -or (Test-Path `$adminCmd)) {
									`$results += 'Running: ' + `$adminCmd + ' ' + `$rstrtSvc + ' ' + `$Process + ' on ' + `$env:COMPUTERNAME
									`$result1 = & `$adminCmd `$rstrtSvc `$Process
									`$results += `$result1 | Out-String
									`$results += 'Running: ' + `$adminCmd + ' ' + `$rstrtSvc + ' ' + `$Process + ' ' + `$prtCmd + ' ' + `$prtTls + ' on ' + `$env:COMPUTERNAME
									`$result2 = & `$adminCmd `$rstrtSvc `$Process `$prtCmd `$prtTls
									`$results += `$result2 | Out-String
									`$commandExecuted = `$true
									break
								}
							} catch {
								`$results += 'Failed to execute ' + `$adminCmd + ' for ' + `$Process + ' on ' + `$env:COMPUTERNAME + ' - ' + `$_
							}
						}

						if (-not `$commandExecuted) {
							throw 'AdminCommand.exe not found in any specified location on ' + `$env:COMPUTERNAME
						}

						return `$results | Out-String
					}"
					Execute-CommandInNewWindow -Server $server -Command $command -WindowTitle "PROCESSES $server - RESTART_SERVICE $process" -Delay $delay
					$delay += 2  # 2 second delay between each server/process combination
				}
			}
			Log-Message "=== PROCESSES RESTART_SERVICE OPERATION COMPLETED ==="
		})

# ==================== PROCESSES LS SECTION ====================
$btnProcessLS.Add_Click({
    if (-not (Validate-Servers)) { return }
    
    $selectedServers = Get-AllSelectedServers
    $selectedProcesses = $processBox.CheckedItems
    
    if ($selectedProcesses.Count -eq 0) {
        Log-Message "WARNING: No processes selected for LS operation."
        return
    }
    
    Log-Message "=== PROCESSES LS OPERATION STARTED ==="
    Log-Message "Selected Servers: $($selectedServers -join ', ')"
    Log-Message "Selected Processes: $($selectedProcesses -join ', ')"
    
    # Show progress message
    $progressForm = New-Object System.Windows.Forms.Form
    $progressForm.Text = "Gathering Process List..."
    $progressForm.Size = New-Object System.Drawing.Size(400, 150)
    $progressForm.StartPosition = "CenterParent"
    $progressForm.FormBorderStyle = "FixedDialog"
    $progressForm.MaximizeBox = $false
    $progressForm.MinimizeBox = $false
    
    # Progress label
    $progressLabel = New-Object System.Windows.Forms.Label
    $progressLabel.Text = "Please wait while gathering process list from servers..."
    $progressLabel.Location = New-Object System.Drawing.Point(20, 30)
    $progressLabel.Size = New-Object System.Drawing.Size(350, 60)
    $progressLabel.TextAlign = "MiddleCenter"
    $progressForm.Controls.Add($progressLabel)
    
    $progressForm.Show()
    $progressForm.Refresh()
    
    # Create findstr filter from selected processes
    $findstrFilter = $selectedProcesses -join ' '
    
    # Build formatted report
    $report = @()
    $report += "=== PROCESSES LS REPORT ==="
    $report += "Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    $report += "Servers: $($selectedServers.Count) selected"
    $report += "Filtered Processes: $($selectedProcesses -join ', ')"
    $report += ""
    
    foreach ($server in $selectedServers | Sort-Object) {
        try {
            $progressLabel.Text = "Gathering process list from: $server"
            $progressForm.Refresh()
            
            Log-Message "Getting process list from $server..."
            
            $result = Invoke-Command -ComputerName $server -ScriptBlock {
                Param ($FilterProcesses)
                $imaginetRoot = [System.Environment]::GetEnvironmentVariable('imaginet_root')
                $adminCmdPaths = if ($imaginetRoot) {
                    @(
                        (Join-Path $imaginetRoot 'utils\AdminCommand.exe'),
                        (Join-Path $imaginetRoot 'Portal\Dll\SVRender\AdminCommand.exe')
                    )
                } else {
                    @('AdminCommand.exe')
                }
                
                $commandExecuted = $false
                $results = @()

                foreach ($adminCmd in $adminCmdPaths) {
                    try {
                        if ($adminCmd -eq 'AdminCommand.exe' -or (Test-Path $adminCmd)) {
                            $results += "=========================================="
                            $results += "SERVER: $env:COMPUTERNAME"
                            $results += "Command: $adminCmd ls -p all | findstr `"$FilterProcesses`""
                            $results += "=========================================="
                            $lsResult = & $adminCmd ls -p all 2>&1
							# Use quotes around the filter for findstr to handle multiple processes
							$filteredResult = $lsResult | findstr "`"$FilterProcesses`""
                            $results += $filteredResult
                            $commandExecuted = $true
                            break
                        }
                    } catch {
                        $results += "Failed to execute $adminCmd ls -p all on $env:COMPUTERNAME - $_"
                    }
                }

                if (-not $commandExecuted) {
                    return "ERROR: AdminCommand.exe not found in any specified location on $env:COMPUTERNAME"
                }

                return $results -join "`n"
            } -ArgumentList $findstrFilter -ErrorAction Stop
            
            $report += $result -split "`n"
            $report += ""
            Log-Message "Successfully retrieved process list from $server"
            
        } catch {
            $report += "ERROR connecting to $server - $_"
            $report += ""
            Log-Message "ERROR connecting to $server - $_"
        }
    }
    
    $progressForm.Close()
    
    # Create results window
    $resultsForm = New-Object System.Windows.Forms.Form
    $resultsForm.Text = "Process List Report - $($selectedServers.Count) Servers, Filtered: $findstrFilter"
    $resultsForm.Size = New-Object System.Drawing.Size(1250, 700)
    $resultsForm.StartPosition = "CenterParent"
    $resultsForm.FormBorderStyle = "Sizable"
    
    # Text box for results
    $resultsTextBox = New-Object System.Windows.Forms.TextBox
    $resultsTextBox.Location = New-Object System.Drawing.Point(10, 10)
    $resultsTextBox.Size = New-Object System.Drawing.Size(1210, 600)
    $resultsTextBox.Multiline = $true
    $resultsTextBox.ScrollBars = "Both"
    $resultsTextBox.Font = New-Object System.Drawing.Font("Consolas", 9)
    $resultsTextBox.WordWrap = $false
    $resultsTextBox.ReadOnly = $true
	$resultsTextBox.Text = $report -join "`r`n"
	$resultsTextBox.SelectionStart = 0
	$resultsTextBox.SelectionLength = 0
    $resultsTextBox.Anchor = "Top, Left, Bottom, Right"
    $resultsForm.Controls.Add($resultsTextBox)
    
    # Copy to Clipboard button
    $copyButton = New-Object System.Windows.Forms.Button
    $copyButton.Text = "Copy to Clipboard"
    $copyButton.Location = New-Object System.Drawing.Point(350, 620)
    $copyButton.Size = New-Object System.Drawing.Size(120, 30)
    $copyButton.Anchor = "Bottom"
    $copyButton.BackColor = [System.Drawing.Color]::LightGreen
    $copyButton.Font = New-Object System.Drawing.Font("Arial", 9, [System.Drawing.FontStyle]::Bold)
    $copyButton.Add_Click({
        try {
            [System.Windows.Forms.Clipboard]::SetText($resultsTextBox.Text)
            $copyButton.Text = "Copied!"
            $copyButton.BackColor = [System.Drawing.Color]::Gold
            $timer = New-Object System.Windows.Forms.Timer
            $timer.Interval = 1500
            $timer.Add_Tick({
                $copyButton.Text = "Copy to Clipboard"
                $copyButton.BackColor = [System.Drawing.Color]::LightGreen
                $timer.Stop()
                $timer.Dispose()
            })
            $timer.Start()
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Failed to copy to clipboard: $_", "Copy Error", "OK", "Error")
        }
    })
    $resultsForm.Controls.Add($copyButton)
    
    # Close button
    $closeButton = New-Object System.Windows.Forms.Button
    $closeButton.Text = "Close"
    $closeButton.Location = New-Object System.Drawing.Point(480, 620)
    $closeButton.Size = New-Object System.Drawing.Size(80, 30)
    $closeButton.Anchor = "Bottom"
    $closeButton.Add_Click({
        $resultsForm.Close()
    })
    $resultsForm.Controls.Add($closeButton)
    
    # Show results window
    $resultsForm.ShowDialog()
    
    Log-Message "=== PROCESSES LS OPERATION COMPLETED ==="
})

# ==================== PROCESSES STATUS SECTION ====================
$btnProcessStatus.Add_Click({
    if (-not (Validate-Servers)) { return }
    
    $selectedServers = Get-AllSelectedServers
    
    Log-Message "=== PROCESSES STATUS OPERATION STARTED ==="
    Log-Message "Selected Servers: $($selectedServers -join ', ')"
    
    # Show progress message
    $progressForm = New-Object System.Windows.Forms.Form
    $progressForm.Text = "Gathering Process Status..."
    $progressForm.Size = New-Object System.Drawing.Size(400, 150)
    $progressForm.StartPosition = "CenterParent"
    $progressForm.FormBorderStyle = "FixedDialog"
    $progressForm.MaximizeBox = $false
    $progressForm.MinimizeBox = $false
    
    # Progress label
    $progressLabel = New-Object System.Windows.Forms.Label
    $progressLabel.Text = "Please wait while gathering process status from servers..."
    $progressLabel.Location = New-Object System.Drawing.Point(20, 30)
    $progressLabel.Size = New-Object System.Drawing.Size(350, 60)
    $progressLabel.TextAlign = "MiddleCenter"
    $progressForm.Controls.Add($progressLabel)
    
    $progressForm.Show()
    $progressForm.Refresh()
    
    # Build formatted report and capture per-server details
    $report = @()
    $report += "=== PROCESSES STATUS REPORT ==="
    $report += "Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    $report += "Servers: $($selectedServers.Count) selected"
    $report += ""
    $perServer = @{}
    
    # For each server, get the process status
    foreach ($server in $selectedServers | Sort-Object) {
        try {
            $progressLabel.Text = "Gathering process status from: $server"
            $progressForm.Refresh()
            
            Log-Message "Getting process status from $server..."
            
            $result = Invoke-Command -ComputerName $server -ScriptBlock {
                $imaginetRoot = [System.Environment]::GetEnvironmentVariable('imaginet_root')
                $adminCmdPaths = if ($imaginetRoot) {
                    @(
                        (Join-Path $imaginetRoot 'utils\AdminCommand.exe'),
                        (Join-Path $imaginetRoot 'Portal\Dll\SVRender\AdminCommand.exe')
                    )
                } else {
                    @('AdminCommand.exe')
                }
                
                $commandExecuted = $false
                $results = @()

                foreach ($adminCmd in $adminCmdPaths) {
                    try {
                        if ($adminCmd -eq 'AdminCommand.exe' -or (Test-Path $adminCmd)) {
                            $results += "=========================================="
                            $results += "SERVER: $env:COMPUTERNAME"
                            $results += "Command: $adminCmd status"
                            $results += "=========================================="
                            $result = & $adminCmd status 2>&1
                            $results += $result
                            $commandExecuted = $true
                            break
                        }
                    } catch {
                        $results += "Failed to execute $adminCmd status on $env:COMPUTERNAME - $_"
                    }
                }

                if (-not $commandExecuted) {
                    return "ERROR: AdminCommand.exe not found in any specified location on $env:COMPUTERNAME"
                }

                return $results -join "`n"
            } -ErrorAction Stop
            
            $perServer[$server] = $result
            $report += $result -split "`n"
            $report += ""
            Log-Message "Successfully retrieved process status from $server"
            
        } catch {
            $report += "ERROR connecting to $server - $_"
            $report += ""
            Log-Message "ERROR connecting to $server - $_"
        }
    }
    
    $progressForm.Close()
    
    # For each server, open a separate window with a parsed table
    foreach ($server in ($selectedServers | Sort-Object)) {
        $text = $perServer[$server]
        if (-not $text) { continue }

        # Parse AdminCommand status lines into table rows
        $rows = @()
        $lines = $text -split "`n"
        foreach ($line in $lines) {
            $l = $line.Trim()
            if (-not $l) { continue }
            $m = [regex]::Match($l, '^\((?<svc>[^)]+)\):\s*Max Services:\s*(?<Max>\d+),\s*Pooled Services:\s*(?<Pooled>\d+),\s*Raised Services:\s*(?<Raised>\d+),\s*Idle Services:\s*(?<Idle>\d+),\s*Active Services:\s*(?<Active>\d+),\s*Pending Requests:\s*(?<PendingReq>\d+),\s*Total Requests:\s*(?<TotalReq>\d+)')
            if ($m.Success) {
                $rows += [PSCustomObject]@{
                    SERVICE    = $m.Groups['svc'].Value
                    Max        = [int]$m.Groups['Max'].Value
                    Pooled     = [int]$m.Groups['Pooled'].Value
                    Raised     = [int]$m.Groups['Raised'].Value
                    Idle       = [int]$m.Groups['Idle'].Value
                    Active     = [int]$m.Groups['Active'].Value
                    PendingReq = [int]$m.Groups['PendingReq'].Value
                    TotalReq   = [int]$m.Groups['TotalReq'].Value
                }
            }
        }
        # Create form per server
        $dc = ([regex]::Match($server, '(\d{4})$')).Value
        if (-not $dc) { $dc = $server.Substring([Math]::Max(0, $server.Length - 4)) }
        $formSrv = New-Object System.Windows.Forms.Form
        $formSrv.Text = "Process Status - $server (DC $dc)"
        $formSrv.Size = New-Object System.Drawing.Size(1100, 700)
        $formSrv.StartPosition = "CenterParent"
        $formSrv.FormBorderStyle = "Sizable"
    $formSrv.KeyPreview = $true
    $formSrv.Add_KeyDown({ if ($_.KeyCode -eq 'Escape') { $formSrv.Close() } })

        $lblHdr = New-Object System.Windows.Forms.Label
        $lblHdr.Text = "SERVICE status on $server (DC $dc)"
        $lblHdr.Font = New-Object System.Drawing.Font("Arial", 11, [System.Drawing.FontStyle]::Bold)
        $lblHdr.Location = New-Object System.Drawing.Point(10, 10)
        $lblHdr.Size = New-Object System.Drawing.Size(800, 22)
        $formSrv.Controls.Add($lblHdr)

        $lvSrv = New-Object System.Windows.Forms.ListView
        $lvSrv.View = 'Details'
        $lvSrv.FullRowSelect = $true
        $lvSrv.GridLines = $true
        $lvSrv.Location = New-Object System.Drawing.Point(10, 40)
        $lvSrv.Size = New-Object System.Drawing.Size(1060, 560)
        $lvSrv.Anchor = "Top, Left, Bottom, Right"
        [void]$lvSrv.Columns.Add("SERVICE", 260)
        [void]$lvSrv.Columns.Add("Max", 60)
        [void]$lvSrv.Columns.Add("Pooled", 70)
        [void]$lvSrv.Columns.Add("Raised", 70)
        [void]$lvSrv.Columns.Add("Idle", 60)
        [void]$lvSrv.Columns.Add("Active", 70)
        [void]$lvSrv.Columns.Add("PendingReq", 90)
        [void]$lvSrv.Columns.Add("TotalReq", 100)


        foreach ($r in ($rows | Sort-Object SERVICE)) {
            $item = New-Object System.Windows.Forms.ListViewItem($r.SERVICE)
            [void]$item.SubItems.Add([string]$r.Max)
            [void]$item.SubItems.Add([string]$r.Pooled)
            [void]$item.SubItems.Add([string]$r.Raised)
            [void]$item.SubItems.Add([string]$r.Idle)
            [void]$item.SubItems.Add([string]$r.Active)
            [void]$item.SubItems.Add([string]$r.PendingReq)
            [void]$item.SubItems.Add([string]$r.TotalReq)
            if ($r.PendingReq -gt 0) {
                # Light red background for any row with PendingReq > 0
                $item.BackColor = [System.Drawing.Color]::FromArgb(255, 235, 238)  # #FFEBEE
            }
            [void]$lvSrv.Items.Add($item)
        }
        $formSrv.Controls.Add($lvSrv)

        $btnCopy = New-Object System.Windows.Forms.Button
        $btnCopy.Text = "Copy to Clipboard"
        $btnCopy.Location = New-Object System.Drawing.Point(10, 610)
        $btnCopy.Size = New-Object System.Drawing.Size(130, 30)
        $btnCopy.Anchor = "Bottom, Left"
        $btnCopy.BackColor = [System.Drawing.Color]::LightGreen
        $btnCopy.Font = New-Object System.Drawing.Font("Arial", 9, [System.Drawing.FontStyle]::Bold)
        $btnCopy.Add_Click({
            try {
                $sb = New-Object System.Text.StringBuilder
                [void]$sb.AppendLine("SERVICE\tMax\tPooled\tRaised\tIdle\tActive\tPendingReq\tTotalReq")
                foreach ($it in $lvSrv.Items) {
                    $cols = @()
                    $cols += $it.Text
                    for ($i=1; $i -lt $it.SubItems.Count; $i++) { $cols += $it.SubItems[$i].Text }
                    [void]$sb.AppendLine(($cols -join "`t"))
                }
                [System.Windows.Forms.Clipboard]::SetText($sb.ToString())
                $btnCopy.Text = "Copied!"
                $btnCopy.BackColor = [System.Drawing.Color]::Gold
                $timer = New-Object System.Windows.Forms.Timer
                $timer.Interval = 1500
                $timer.Add_Tick({
                    $btnCopy.Text = "Copy to Clipboard"
                    $btnCopy.BackColor = [System.Drawing.Color]::LightGreen
                    $timer.Stop()
                    $timer.Dispose()
                })
                $timer.Start()
            } catch {
                [System.Windows.Forms.MessageBox]::Show("Failed to copy to clipboard: $_", "Copy Error", "OK", "Error")
            }
        })
        $formSrv.Controls.Add($btnCopy)

        $btnClose = New-Object System.Windows.Forms.Button
        $btnClose.Text = "Close"
        $btnClose.Location = New-Object System.Drawing.Point(150, 610)
        $btnClose.Size = New-Object System.Drawing.Size(90, 30)
        $btnClose.Anchor = "Bottom, Left"
        $btnClose.Add_Click({ $formSrv.Close() })
        $formSrv.Controls.Add($btnClose)

        # Show non-modal so multiple servers can be reviewed
        $formSrv.Show()
    }
    
    Log-Message "=== PROCESSES STATUS OPERATION COMPLETED ==="
})

# ==================== IIS APP POOLS RECYCLE SECTION ====================
$btnIISRecycle.Add_Click({
    if (-not (Validate-Servers)) { return }
    
    $selectedServers = Get-AllSelectedServers
    $selectedAppPools = $poolBox.CheckedItems
    
    if ($selectedAppPools.Count -eq 0) {
        Log-Message "WARNING: No IIS App Pools selected for recycle."
        return
    }
    
    Log-Message "=== IIS APP POOLS RECYCLE OPERATION STARTED ==="
    Log-Message "Selected Servers: $($selectedServers -join ', ')"
    Log-Message "Selected IIS App Pools: $($selectedAppPools -join ', ')"
    
    foreach ($Server in $selectedServers) {
        Log-Message "Processing IIS App Pools on ${Server}..."
        
        foreach ($AppPool in $selectedAppPools) {
            try {
                Log-Message "Running: Restart-WebAppPool -Name $AppPool on $Server"
                $result = Invoke-Command -ComputerName $Server -ScriptBlock {
                    Param ($AppPool)
                    Import-Module WebAdministration
                    if (Test-Path "IIS:\AppPools\$AppPool") {
                        Restart-WebAppPool -Name $AppPool -ErrorAction Stop
                        return "${AppPool} recycled successfully on $env:COMPUTERNAME."
                    } else {
                        return "IIS App Pool ${AppPool} not found on $env:COMPUTERNAME."
                    }
                } -ArgumentList $AppPool -ErrorAction Stop
                Log-Message $result
            } catch {
                Log-Message "ERROR recycling IIS App Pool ${AppPool} on ${Server} - $_"
            }
        }
        Log-Message "IIS App Pools recycle processing completed for ${Server}"
    }
    Log-Message "=== IIS APP POOLS RECYCLE OPERATION COMPLETED ==="
})

## Removed IIS APP POOLS HOLY3 SECTION (handled via COMMON TASKS)

# ==================== IIS APP POOLS START SECTION ====================
$btnIISStart.Add_Click({
    if (-not (Validate-Servers)) { return }
    
    $selectedServers = Get-AllSelectedServers
    $selectedAppPools = $poolBox.CheckedItems
    
    if ($selectedAppPools.Count -eq 0) {
        Log-Message "WARNING: No IIS App Pools selected for start."
        return
    }
    
    Log-Message "=== IIS APP POOLS START OPERATION STARTED ==="
    Log-Message "Selected Servers: $($selectedServers -join ', ')"
    Log-Message "Selected IIS App Pools: $($selectedAppPools -join ', ')"
    
    foreach ($Server in $selectedServers) {
        Log-Message "Starting IIS App Pools on ${Server}..."
        
        foreach ($AppPool in $selectedAppPools) {
            try {
                Log-Message "Running: Start-WebAppPool -Name $AppPool on $Server"
                $result = Invoke-Command -ComputerName $Server -ScriptBlock {
                    Param ($AppPool)
                    Import-Module WebAdministration
                    if (Test-Path "IIS:\AppPools\$AppPool") {
                        Restart-WebAppPool -Name $AppPool -ErrorAction Stop
                        return "[$env:COMPUTERNAME] ${AppPool} recycled successfully."
                    } else {
                        return "[$env:COMPUTERNAME] IIS App Pool ${AppPool} not found."
                    }
                } -ArgumentList $AppPool -ErrorAction Stop
                Log-Message $result
            } catch {
                Log-Message "ERROR starting IIS App Pool ${AppPool} on ${Server} - $_"
            }
        }
        Log-Message "IIS App Pools start processing completed for ${Server}"
    }
    Log-Message "=== IIS APP POOLS START OPERATION COMPLETED ==="
})

# ==================== IIS APP POOLS STOP SECTION ====================
$btnIISStop.Add_Click({
    if (-not (Validate-Servers)) { return }
    
    $selectedServers = Get-AllSelectedServers
    $selectedAppPools = $poolBox.CheckedItems
    
    if ($selectedAppPools.Count -eq 0) {
        Log-Message "WARNING: No IIS App Pools selected for stop."
        return
    }
    
    Log-Message "=== IIS APP POOLS STOP OPERATION STARTED ==="
    Log-Message "Selected Servers: $($selectedServers -join ', ')"
    Log-Message "Selected IIS App Pools: $($selectedAppPools -join ', ')"
    
    foreach ($Server in $selectedServers) {
        Log-Message "Stopping IIS App Pools on ${Server}..."
        
        foreach ($AppPool in $selectedAppPools) {
            try {
                Log-Message "Running: Stop-WebAppPool -Name $AppPool on $Server"
                $result = Invoke-Command -ComputerName $Server -ScriptBlock {
                    Param ($AppPool)
                    Import-Module WebAdministration
                    if (Test-Path "IIS:\AppPools\$AppPool") {
                        Stop-WebAppPool -Name $AppPool -ErrorAction Stop
                        return "${AppPool} stopped successfully on $env:COMPUTERNAME."
                    } else {
                        return "IIS App Pool ${AppPool} not found on $env:COMPUTERNAME."
                    }
                } -ArgumentList $AppPool -ErrorAction Stop
                Log-Message $result
            } catch {
                Log-Message "ERROR stopping IIS App Pool ${AppPool} on ${Server} - $_"
            }
        }
        Log-Message "IIS App Pools stop processing completed for ${Server}"
    }
    Log-Message "=== IIS APP POOLS STOP OPERATION COMPLETED ==="
})

# ==================== MAIN CONTROL BUTTONS ====================
# UNSELECT ALL Button
$unselectAllButton = New-Object System.Windows.Forms.Button
$unselectAllButton.Text = "UNSELECT`nALL"
$unselectAllButton.Location = New-Object System.Drawing.Point(10, 770)
$unselectAllButton.Size = New-Object System.Drawing.Size(100, 38)
$unselectAllButton.Anchor = "Bottom"
$unselectAllButton.BackColor = [System.Drawing.Color]::FromArgb(226, 239, 217)
$unselectAllButton.Font = New-Object System.Drawing.Font("Arial", 9, [System.Drawing.FontStyle]::Bold)
$unselectAllButton.Add_Click({
    # Unselect all servers
    for ($i = 0; $i -lt $coreServerBox.Items.Count; $i++) {
        $coreServerBox.SetItemChecked($i, $false)
    }
    for ($i = 0; $i -lt $daServerBox.Items.Count; $i++) {
        $daServerBox.SetItemChecked($i, $false)
    }
    for ($i = 0; $i -lt $portalServerBox.Items.Count; $i++) {
        $portalServerBox.SetItemChecked($i, $false)
    }
    for ($i = 0; $i -lt $otherServerBox.Items.Count; $i++) {
        $otherServerBox.SetItemChecked($i, $false)
    }
    
    # Unselect all services
    for ($i = 0; $i -lt $serviceBox.Items.Count; $i++) {
        $serviceBox.SetItemChecked($i, $false)
    }
    
    # Unselect all processes
    for ($i = 0; $i -lt $processBox.Items.Count; $i++) {
        $processBox.SetItemChecked($i, $false)
    }
    
    # Unselect all IIS App Pools
    for ($i = 0; $i -lt $poolBox.Items.Count; $i++) {
        $poolBox.SetItemChecked($i, $false)
    }
    
    Log-Message "All selections cleared by user."
})
$form.Controls.Add($unselectAllButton)

# Consolidated Edit Button
$editButton = New-Object System.Windows.Forms.Button
$editButton.Text = "Edit Svr/Svc"
$editButton.Location = New-Object System.Drawing.Point(120, 775)
$editButton.Size = New-Object System.Drawing.Size(110, 30)
$editButton.Anchor = "Bottom"
$editButton.BackColor = [System.Drawing.Color]::LightGray
$editButton.ForeColor = [System.Drawing.Color]::Black
$editButton.Font = New-Object System.Drawing.Font("Arial", 9, [System.Drawing.FontStyle]::Bold)
$editButton.Add_Click({
    if (Show-TabbedEditWindow) {
        Refresh-ListBoxes
    }
})
$form.Controls.Add($editButton)

# Clear Log Button 
$clearLogButton = New-Object System.Windows.Forms.Button
$clearLogButton.Text = "Clear Log"
$clearLogButton.Location = New-Object System.Drawing.Point(233, 775)
$clearLogButton.Size = New-Object System.Drawing.Size(70, 30)
$clearLogButton.Anchor = "Bottom"
$clearLogButton.BackColor = [System.Drawing.Color]::LightGray
$clearLogButton.ForeColor = [System.Drawing.Color]::Black
$clearLogButton.Font = New-Object System.Drawing.Font("Arial", 8, [System.Drawing.FontStyle]::Bold)
$clearLogButton.Add_Click({
    $logBox.Text = ""
    Log-Message "Log cleared by user."
})
$form.Controls.Add($clearLogButton)

# Close All Windows Button 
$closeWindowsButton = New-Object System.Windows.Forms.Button
$closeWindowsButton.Text = "Close All Windows"
$closeWindowsButton.Location = New-Object System.Drawing.Point(308, 775)
$closeWindowsButton.Size = New-Object System.Drawing.Size(130, 30)
$closeWindowsButton.Anchor = "Bottom"
$closeWindowsButton.BackColor = [System.Drawing.Color]::LightGray
$closeWindowsButton.ForeColor = [System.Drawing.Color]::Black
$closeWindowsButton.Font = New-Object System.Drawing.Font("Arial", 8, [System.Drawing.FontStyle]::Bold)
$closeWindowsButton.Add_Click({
    foreach ($window in $global:PowerShellWindows) {
        try {
            if (-not $window.HasExited) { $window.CloseMainWindow() }
        } catch { }
    }
    $global:PowerShellWindows = @()
    Log-Message "All PowerShell windows closed."
})
$form.Controls.Add($closeWindowsButton)

# ==================== RESTART ALL BUTTON ====================
$restartAllButton = New-Object System.Windows.Forms.Button
$restartAllButton.Text = "RESTART ALL"
$restartAllButton.Location = New-Object System.Drawing.Point(1145, 775)
$restartAllButton.Size = New-Object System.Drawing.Size(90, 34)
$restartAllButton.Anchor = "Bottom, Right"
$restartAllButton.BackColor = [System.Drawing.Color]::LightBlue
$restartAllButton.Font = New-Object System.Drawing.Font("Arial", 9, [System.Drawing.FontStyle]::Bold)
$restartAllButton.Add_Click({
    if (-not (Validate-Servers)) { return }

    $selectedServers = Get-AllSelectedServers
    $selectedServices = $serviceBox.CheckedItems
    $selectedProcesses = $processBox.CheckedItems
    $selectedAppPools = $poolBox.CheckedItems

    Log-Message "======================================="
    Log-Message "=== RESTART ALL OPERATION STARTED ==="
    Log-Message "======================================="
    Log-Message "Selected Servers: $($selectedServers -join ', ')"
    Log-Message "Selected Services: $($selectedServices -join ', ')"
    Log-Message "Selected Processes: $($selectedProcesses -join ', ')"
    Log-Message "Selected IIS App Pools: $($selectedAppPools -join ', ')"

    $rstrtSvc = 'restart_service'
    $prtCmd = '-p'
    $prtTls = 'all'

    foreach ($Server in $selectedServers) {
        Log-Message "Processing ${Server}..."

        # ==================== PROCESSES RESTART SECTION ====================
        if ($selectedProcesses.Count -gt 0) {
            Log-Message "--- RESTARTING PROCESSES ON $Server ---"
            foreach ($Process in $selectedProcesses) {
                try {
                    Log-Message "Attempting to restart process: $Process on $Server"
                    $runAD = @{
                        ComputerName = $Server
                        ScriptBlock  = {
                            Param ($Process, $rstrtSvc, $prtCmd, $prtTls)
                            $imaginetRoot = [System.Environment]::GetEnvironmentVariable("imaginet_root")
                            $adminCmdPaths = if ($imaginetRoot) {
                                @(
                                    (Join-Path $imaginetRoot "utils\AdminCommand.exe"),
                                    (Join-Path $imaginetRoot "Portal\Dll\SVRender\AdminCommand.exe")
                                )
                            } else {
                                @("AdminCommand.exe")
                            }
                            
                            $commandExecuted = $false
                            $results = @()

                            foreach ($adminCmd in $adminCmdPaths) {
                                try {
                                    if ($adminCmd -eq "AdminCommand.exe" -or (Test-Path $adminCmd)) {
                                        $results += "Running: $adminCmd $rstrtSvc $Process"
                                        $result1 = & $adminCmd $rstrtSvc $Process
                                        $results += $result1 | Out-String
                                        $results += "Running: $adminCmd $rstrtSvc $Process $prtCmd $prtTls"
                                        $result2 = & $adminCmd $rstrtSvc $Process $prtCmd $prtTls
                                        $results += $result2 | Out-String
                                        $commandExecuted = $true
                                        break
                                    }
                                } catch {
                                    $results += "Failed to execute $adminCmd for $Process on $env:COMPUTERNAME - $_"
                                }
                            }

                            if (-not $commandExecuted) {
                                throw "AdminCommand.exe not found in any specified location on $env:COMPUTERNAME."
                            }

                            return $results | Out-String
                        }
                        ArgumentList = $Process, $rstrtSvc, $prtCmd, $prtTls
                    }
                    $result = Invoke-Command @runAD | Select-String "AdminCommand.log" -CaseSensitive -NotMatch
                    foreach ($line in $result -split "`n") {
                        if ($line.Trim()) {
                            Log-Message $line
                        }
                    }
                    Log-Message "Process ${Process} restarted on ${Server}"
                } catch {
                    Log-Message "ERROR restarting process ${Process} on ${Server} - $_"
                }
            }
        }

        # ==================== SERVICES RESTART SECTION ====================
        if ($selectedServices.Count -gt 0) {
            Log-Message "--- RESTARTING SERVICES ON $Server ---"
            foreach ($Service in $selectedServices) {
                try {
                    Log-Message "Running: Restart-Service -Name $Service"
                    $result = Invoke-Command -ComputerName $Server -ScriptBlock {
                        Param ($Service)
                        if (Get-Service -Name $Service -ErrorAction SilentlyContinue) {
                            Restart-Service -Name $Service -Force -ErrorAction Stop
                            return "${Service} restarted successfully."
                        } else {
                            return "Service ${Service} not found on $env:COMPUTERNAME."
                        }
                    } -ArgumentList $Service -ErrorAction Stop
                    Log-Message $result
                } catch {
                    Log-Message "ERROR restarting service ${Service} on ${Server} - $_"
                }
            }
        }

        # ==================== IIS APP POOLS RESTART SECTION ====================
        if ($selectedAppPools.Count -gt 0) {
            Log-Message "--- RESTARTING IIS APP POOLS ON $Server ---"
            foreach ($AppPool in $selectedAppPools) {
                try {
                    Log-Message "Running: Restart-WebAppPool -Name $AppPool"
                    $result = Invoke-Command -ComputerName $Server -ScriptBlock {
                        Param ($AppPool)
                        Import-Module WebAdministration
                        if (Test-Path "IIS:\AppPools\$AppPool") {
                            Restart-WebAppPool -Name $AppPool -ErrorAction Stop
                            return "${AppPool} restarted successfully."
                        } else {
                            return "IIS App Pool ${AppPool} not found on $env:COMPUTERNAME."
                        }
                    } -ArgumentList $AppPool -ErrorAction Stop
                    Log-Message $result
                } catch {
                    Log-Message "ERROR restarting IIS App Pool ${AppPool} on ${Server} - $_"
                }
            }
        }

        Log-Message "Done with ${Server}"
    }
    Log-Message "======================================="
    Log-Message "=== RESTART ALL OPERATION COMPLETED ==="
    Log-Message "======================================="
})
$form.Controls.Add($restartAllButton)

### COMMON TASKS: functions and dropdown + RUN button

# Shared task functions
function Invoke-DisableMemo {
    if (-not (Validate-Servers)) { return }

    $selectedServers = Get-AllSelectedServers
    $memoServices = @('PhilipsMemoProductExporterWrapperService', 'PhilipsMemoPrometheusService', 'ProductExporterService', 'windows_exporter')

    $confirmResult = [System.Windows.Forms.MessageBox]::Show(
        "Are you sure you want to STOP and DISABLE all Memo services on selected servers?",
        "Confirm DISABLE MEMO Operation",
        "YesNo",
        "Warning"
    )
    if ($confirmResult -eq "No") {
        Log-Message "DISABLE MEMO operation cancelled by user."
        return
    }

    Log-Message "=== DISABLE MEMO OPERATION STARTED ==="
    Log-Message "Selected Servers: $($selectedServers -join ', ')"
    Log-Message "Memo Services: $($memoServices -join ', ')"

    foreach ($Server in $selectedServers) {
        Log-Message "Disabling Memo Services on ${Server}..."
        foreach ($Service in $memoServices) {
            try {
                Log-Message "Running: Stop-Service and Set-Service -StartupType Disabled for $Service on $Server"
                $result = Invoke-Command -ComputerName $Server -ScriptBlock {
                    Param ($Service)
                    if (Get-Service -Name $Service -ErrorAction SilentlyContinue) {
                        try { Stop-Service -Name $Service -Force -ErrorAction Stop; $stopResult = "${Service} stopped successfully" } catch { $stopResult = "Warning: Could not stop ${Service} - $_" }
                        Set-Service -Name $Service -StartupType Disabled -ErrorAction Stop
                        return "$stopResult and ${Service} disabled successfully on $env:COMPUTERNAME."
                    } else {
                        return "Service ${Service} not found on $env:COMPUTERNAME."
                    }
                } -ArgumentList $Service -ErrorAction Stop
                Log-Message $result
            } catch {
                Log-Message "ERROR disabling service ${Service} on ${Server} - $_"
            }
        }
        Log-Message "Memo services disable processing completed for ${Server}"
    }
    Log-Message "=== DISABLE MEMO OPERATION COMPLETED ==="
}

function Invoke-EnableMemo {
    if (-not (Validate-Servers)) { return }

    $selectedServers = Get-AllSelectedServers
    $memoServices = @('PhilipsMemoProductExporterWrapperService', 'PhilipsMemoPrometheusService', 'ProductExporterService', 'windows_exporter')

    Log-Message "=== ENABLE MEMO OPERATION STARTED ==="
    Log-Message "Selected Servers: $($selectedServers -join ', ')"
    Log-Message "Memo Services: $($memoServices -join ', ')"

    foreach ($Server in $selectedServers) {
        Log-Message "Enabling Memo Services on ${Server}..."
        foreach ($Service in $memoServices) {
            try {
                Log-Message "Running: Set-Service -StartupType Automatic and Start-Service for $Service on $Server"
                $result = Invoke-Command -ComputerName $Server -ScriptBlock {
                    Param ($Service)
                    if (Get-Service -Name $Service -ErrorAction SilentlyContinue) {
                        Set-Service -Name $Service -StartupType Automatic -ErrorAction Stop
                        try { Start-Service -Name $Service -ErrorAction Stop; $startResult = "${Service} started successfully" } catch { $startResult = "Warning: Could not start ${Service} - $_" }
                        return "${Service} enabled successfully and $startResult on $env:COMPUTERNAME."
                    } else {
                        return "Service ${Service} not found on $env:COMPUTERNAME."
                    }
                } -ArgumentList $Service -ErrorAction Stop
                Log-Message $result
            } catch {
                Log-Message "ERROR enabling service ${Service} on ${Server} - $_"
            }
        }
        Log-Message "Memo services enable processing completed for ${Server}"
    }
    Log-Message "=== ENABLE MEMO OPERATION COMPLETED ==="
}

function Invoke-OrchestratorNormalizationSvcRestart {
    if (-not (Validate-Servers)) { return }

    $selectedServers = Get-AllSelectedServers
    $appPools = @(
        'WFMOtherQueryHandlersAppPool',
        'WFMWiManagmentAppPool',
        'OrchestratorServiceAppPool',
        'NormalizationAppPool'
    )
    $services = @(
        'Imaginet WFM Scheduled Queries Dispatcher',
        'Imaginet MST Normalizer'
    )

    Log-Message '=== ORCHESTRATOR_NORMALIZATION SVC RESTART STARTED ==='
    Log-Message "Selected Servers: $($selectedServers -join ', ')"
    Log-Message "App Pools (recycle, in order): $($appPools -join ', ')"
    Log-Message "Services (restart, in order): $($services -join ', ')"

    foreach ($Server in $selectedServers) {
        $steps = @()
        foreach ($p in $appPools)  { $steps += [pscustomobject]@{ Type = 'AppPool'; Name = $p } }
        foreach ($s in $services)  { $steps += [pscustomobject]@{ Type = 'Service'; Name = $s } }

        Log-Message "Executing steps on ${Server} (with 3s pauses)..."
        for ($i = 0; $i -lt $steps.Count; $i++) {
            $step = $steps[$i]
            if ($step.Type -eq 'AppPool') {
                try {
                    $result = Invoke-Command -ComputerName $Server -ScriptBlock {
                        Param ($AppPool)
                        Import-Module WebAdministration
                        if (Test-Path "IIS:\\AppPools\\$AppPool") {
                            Restart-WebAppPool -Name $AppPool -ErrorAction Stop
                            return "[$env:COMPUTERNAME] ${AppPool} recycled successfully."
                        } else {
                            return "[$env:COMPUTERNAME] IIS App Pool ${AppPool} not found."
                        }
                    } -ArgumentList $step.Name -ErrorAction Stop
                    Log-Message $result
                } catch {
                    Log-Message "ERROR recycling IIS App Pool ${($step.Name)} on ${Server} - $_"
                }
            } else {
                try {
                    $result = Invoke-Command -ComputerName $Server -ScriptBlock {
                        Param ($Service)
                        $svc = Get-Service -Name $Service -ErrorAction SilentlyContinue
                        if ($null -eq $svc) { return "[$env:COMPUTERNAME] Service '$Service' not found." }
                        try {
                            Restart-Service -Name $Service -Force -ErrorAction Stop
                            return "[$env:COMPUTERNAME] Restarted '$Service'."
                        }
                        catch { return "[$env:COMPUTERNAME] Failed to restart '$Service' - $_" }
                    } -ArgumentList $step.Name -ErrorAction Continue
                    Log-Message $result
                } catch {
                    Log-Message "ERROR restarting service '${($step.Name)}' on ${Server} - $_"
                }
            }

            if ($i -lt ($steps.Count - 1)) { Start-Sleep -Seconds 3 }
        }

        Log-Message "Completed Orchestrator_Normalization Svc Restart for ${Server}"
    }

    Log-Message '=== ORCHESTRATOR_NORMALIZATION SVC RESTART COMPLETED ==='
}

# Build preview text for Common Tasks showing explicit commands (no Invoke-Command wrappers)
function Build-CommonTaskPreview {
    param(
        [string]$TaskName,
        [string[]]$Servers
    )

    $lines = @()
    $lines += "=== REVIEW TASK PREVIEW ==="
    $lines += "Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    $lines += "Selected Task: $TaskName"
    $lines += "Servers: $($Servers -join ', ')"
    $lines += ""

    switch ($TaskName) {
        'AdminCommand status' {
            foreach ($s in $Servers) {
                $lines += "--- $s ---"
                $lines += "AdminCommand.exe status"
                $lines += ""
            }
        }
        'AdminTools Not Working' {
            $appPools = @('AdminServerAppPool','PACSAppPool','PACSAuthenticatedAppPool')
            $svc = 'Tomcat9'
            foreach ($s in $Servers) {
                $lines += "--- $s ---"
                foreach ($p in $appPools) { $lines += "Import-Module WebAdministration; Restart-WebAppPool -Name '$p'" }
                $lines += "if ((Get-Service -Name '$svc').Status -eq 'Running') { Stop-Service -Name '$svc' -Force }"
                $lines += "Start-Service -Name '$svc'"
                $lines += ""
            }
        }
        'Kick it - HARD' {
            $services = @(
                'Imaginet Auto-Router Scheduling Module','Imaginet Task Scanner','Imaginet Task Dispatcher',
                'Imaginet WFM Scheduled Queries Dispatcher','Imaginet MstSync Server','Imaginet Medilink Sync Listener',
                'Imaginet WCF','Imaginet RisSync Server','Imaginet Non-Dicom Auto Ingestion Service',
                'Imaginet Medilink Converter','MSTNormalizer','Imaginet Loader Server',
                'Imaginet IOCM Probe','Imaginet Auto-Router Execution Module'
            )
            foreach ($s in $Servers) {
                $lines += "--- $s ---"
                foreach ($svc in $services) { $lines += "Stop-Service -Name '$svc' -Force" }
                foreach ($svc in $services) { $lines += "Start-Service -Name '$svc'" }
                $lines += ""
            }
        }
        'Orchestrator_Normalization Svc Restart' {
            $appPools = @(
                'WFMOtherQueryHandlersAppPool','WFMWiManagmentAppPool',
                'OrchestratorServiceAppPool','NormalizationAppPool'
            )
            $services = @(
                'Imaginet WFM Scheduled Queries Dispatcher','Imaginet MST Normalizer'
            )
            foreach ($s in $Servers) {
                $lines += "--- $s ---"
                $ordered = @()
                foreach ($p in $appPools) { $ordered += "Import-Module WebAdministration; Restart-WebAppPool -Name '$p'" }
                foreach ($svc in $services) { $ordered += "Restart-Service -Name '$svc' -Force" }
                for ($i = 0; $i -lt $ordered.Count; $i++) {
                    $lines += $ordered[$i]
                    if ($i -lt ($ordered.Count - 1)) { $lines += "Start-Sleep -Seconds 3" }
                }
                $lines += ""
            }
        }
        'IIS Recycle Settings Report' {
            $lines += "Note: This task gathers settings via WebAdministration but does not execute changes."
            foreach ($s in $Servers) {
                $lines += "--- $s ---"
@'
Import-Module WebAdministration
Get-ChildItem IIS:\AppPools | Select-Object Name,
    @{Name="PrivateMemoryKB";Expression={$_.recycling.periodicRestart.privateMemory}},
    @{Name="Requests";Expression={$_.recycling.periodicRestart.requests}},
    @{Name="SpecificTime";Expression={(($_.recycling.periodicRestart.schedule.collection | ForEach-Object { $_.value }) -join ',')}},
    @{Name="TimeHours";Expression={$_.recycling.periodicRestart.time.TotalHours}}
'@ | ForEach-Object { $lines += $_ }
                $lines += ""
            }
        }
        'CSVER REPORT' {
            foreach ($s in $Servers) { $lines += "--- $s ---"; $lines += "csver.py"; $lines += "" }
        }
        'DICOM Node - Apply Changes (DAs)' {
            $processes = @('svdser','svdtc','svqe','svsm','svdsk','svfolder','svdds','svldr','svfwd')
            $svc = 'Imaginet Loader Server'
            foreach ($s in $Servers) {
                $lines += "--- $s ---"
                foreach ($p in $processes) { $lines += "AdminCommand.exe restart_service $p" }
                $lines += "if ((Get-Service -Name '$svc').Status -eq 'Running') { Stop-Service -Name '$svc' -Force }"
                $lines += "Start-Service -Name '$svc'"
                $lines += ""
            }
        }
        'DICOM Node - Make Visible VMotion (Portals)' {
            $pools = @('ArchiveDataAppPool','ArchiveViewsAppPool','UIPatientsAppPool')
            foreach ($s in $Servers) { $lines += "--- $s ---"; foreach ($p in $pools) { $lines += "Import-Module WebAdministration; Restart-WebAppPool -Name '$p'" }; $lines += "" }
        }
        'Disable MEMO' {
            $services = @('PhilipsMemoProductExporterWrapperService','PhilipsMemoPrometheusService','ProductExporterService','windows_exporter')
            foreach ($s in $Servers) { $lines += "--- $s ---"; foreach ($svc in $services) { $lines += "Stop-Service -Name '$svc' -Force"; $lines += "Set-Service -Name '$svc' -StartupType Disabled" }; $lines += "" }
        }
        'Enable MEMO' {
            $services = @('PhilipsMemoProductExporterWrapperService','PhilipsMemoPrometheusService','ProductExporterService','windows_exporter')
            foreach ($s in $Servers) { $lines += "--- $s ---"; foreach ($svc in $services) { $lines += "Set-Service -Name '$svc' -StartupType Automatic"; $lines += "Start-Service -Name '$svc'" }; $lines += "" }
        }
        'FlexLM - Temp License Reset' {
            $lines += 'This script will query/log, then change the Registry Value associated with the FlexLM license to extend the use of a TEMP license.'
            $lines += 'Registry Commands:'
            $lines += '  Get-ItemProperty -Path "HKLM:\SOFTWARE\WOW6432Node\Installer\Products\4B72950E55AF99247B1C6D45B9132807"'
            $lines += '  Set-ItemProperty -Path "HKLM:\SOFTWARE\WOW6432Node\Installer\Products\4B72950E55AF99247B1C6D45B9132807" -Name "Assignment" -Value 1'
            $lines += "  Restart-Service -Name 'FLEXlm Service'"
            foreach ($s in $Servers) {
                $lines += "--- $s ---"
                $lines += '(Get-ItemProperty -Path "HKLM:\SOFTWARE\WOW6432Node\Installer\Products\4B72950E55AF99247B1C6D45B9132807").Assignment'
                $lines += 'Set-ItemProperty -Path "HKLM:\SOFTWARE\WOW6432Node\Installer\Products\4B72950E55AF99247B1C6D45B9132807" -Name "Assignment" -Value 1'
                $lines += "Restart-Service -Name 'FLEXlm Service' -Force"
                $lines += ""
            }
        }
        'Holy3' {
            $pools = @('ArchiveAppSvcAppPool','ArchiveViewsAppPool','ArchiveDataAppPool')
            foreach ($s in $Servers) { $lines += "--- $s ---"; foreach ($p in $pools) { $lines += "Import-Module WebAdministration; Restart-WebAppPool -Name '$p'" }; $lines += "" }
        }
        'Java Service Info' {
                        foreach ($s in $Servers) {
                                $lines += "--- $s ---"
@'
Get-CimInstance Win32_Process |
    Where-Object { $_.Name -like "*java*" } |
    Select-Object ProcessId, Name,
        @{Name="Memory(MB)";Expression={[math]::Round($_.WorkingSetSize/1MB,2)}},
        @{Name="AppName";Expression={ if($_.CommandLine -match '-Dapp.name=([^\s]+)'){ $matches[1] } else { "N/A" } }} |
    Sort-Object AppName
'@ | ForEach-Object { $lines += $_ }
                                $lines += ""
                        }
        }
        'UPTIME REPORT' {
                        foreach ($s in $Servers) {
                                $lines += "--- $s ---"
@'
Get-CimInstance Win32_OperatingSystem |
    Select-Object CSName,
        @{Name="LastBootupTime";Expression={$_.LastBootUpTime}},
        @{Name="UptimeDays";Expression={[math]::Round((New-TimeSpan -Start $_.LastBootUpTime -End (Get-Date)).TotalDays,2)}}
'@ | ForEach-Object { $lines += $_ }
                                $lines += ""
                        }
        }
        default {
            $lines += "No preview available for task: $TaskName"
        }
    }

    return $lines -join "`r`n"
}

function Invoke-AdminToolsNotWorking {
    if (-not (Validate-Servers)) { return }
    $selectedServers = Get-AllSelectedServers
    $appPools = @('AdminServerAppPool','PACSAppPool','PACSAuthenticatedAppPool')
    $serviceName = 'Tomcat9'

    Log-Message "=== ADMINTools NOT WORKING - RECYCLE APPOOLS AND RESTART SERVICE STARTED ==="
    Log-Message "Selected Servers: $($selectedServers -join ', ')"
    Log-Message "App Pools: $($appPools -join ', ')"
    Log-Message "Windows Service: $serviceName"

    foreach ($Server in $selectedServers) {
        Log-Message "Processing AdminTools fix on ${Server}..."
        # Recycle specified app pools
        foreach ($AppPool in $appPools) {
            try {
                $result = Invoke-Command -ComputerName $Server -ScriptBlock {
                    Param ($AppPool)
                    Import-Module WebAdministration
                    if (Test-Path "IIS:\AppPools\$AppPool") {
                        Restart-WebAppPool -Name $AppPool -ErrorAction Stop
                        return "$AppPool recycled successfully on $env:COMPUTERNAME."
                    } else { return "IIS App Pool $AppPool not found on $env:COMPUTERNAME." }
                } -ArgumentList $AppPool -ErrorAction Stop
                Log-Message $result
            } catch { Log-Message "ERROR recycling IIS App Pool ${AppPool} on ${Server} - $_" }
        }

        # Restart the Windows service
        try {
            $svcResult = Invoke-Command -ComputerName $Server -ScriptBlock {
                Param ($SvcName)
                $svc = Get-Service -Name $SvcName -ErrorAction SilentlyContinue
                if ($null -ne $svc) {
                    if ($svc.Status -eq 'Running') { Stop-Service -Name $SvcName -Force -ErrorAction SilentlyContinue }
                    Start-Sleep -Seconds 1
                    Start-Service -Name $SvcName -ErrorAction Stop
                    return "Service '$SvcName' restarted successfully on $env:COMPUTERNAME."
                } else { return "Service '$SvcName' not found on $env:COMPUTERNAME." }
            } -ArgumentList $serviceName -ErrorAction Stop
            Log-Message $svcResult
        } catch { Log-Message "ERROR restarting service '$serviceName' on ${Server} - $_" }

        Log-Message "AdminTools fix completed for ${Server}"
    }
    Log-Message "=== ADMINTools NOT WORKING TASK COMPLETED ==="
}

function Invoke-FlexLMTempLicenseReset {
    if (-not (Validate-Servers)) { return }

    $selectedServers = Get-AllSelectedServers
    $regPath = 'HKLM:SOFTWARE\WOW6432Node\Installer\Products\4B72950E55AF99247B1C6D45B9132807'
    $regName = 'Assignment'
    $serviceName = 'FLEXlm Service'

    # Preliminary warning note
    $note = @(
        'This script will query/log, then change the Registry Value associated with the FlexLM license to extend the use of a TEMP license.',
        'Registry Commands:',
        '  Get-ItemProperty -Path "HKLM:\\SOFTWARE\\WOW6432Node\\Installer\\Products\\4B72950E55AF99247B1C6D45B9132807"',
        '  Set-ItemProperty -Path "HKLM:\\SOFTWARE\\WOW6432Node\\Installer\\Products\\4B72950E55AF99247B1C6D45B9132807" -Name "Assignment" -Value 1'
    ) -join "`n"
    Log-Message $note

    Log-Message '=== FLEXLM TEMP LICENSE RESET STARTED ==='
    Log-Message ("Selected Servers: {0}" -f ($selectedServers -join ', '))

    foreach ($Server in $selectedServers) {
        Log-Message "Processing FlexLM license reset on ${Server}..."
        try {
            $result = Invoke-Command -ComputerName $Server -ScriptBlock {
                param($p,$n,$svc)
                $out = @()
                try {
                    $cur = Get-ItemProperty -Path $p -ErrorAction Stop
                    $val = if ($null -ne $cur.$n) { $cur.$n } else { '<not set>' }
                    $out += "Current ${n}: $val"
                } catch {
                    $out += "ERROR reading $p - $_"
                }
                try {
                    Set-ItemProperty -Path $p -Name $n -Value 1 -ErrorAction Stop
                    $out += "Set $n to 1"
                } catch {
                    $out += "ERROR setting $n - $_"
                }
                try {
                    if (Get-Service -Name $svc -ErrorAction SilentlyContinue) {
                        Restart-Service -Name $svc -Force -ErrorAction Stop
                        $out += "Restarted service '$svc'"
                    } else {
                        $out += "Service '$svc' not found"
                    }
                } catch {
                    $out += "ERROR restarting service '$svc' - $_"
                }
                return $out -join "`n"
            } -ArgumentList $regPath,$regName,$serviceName -ErrorAction Stop
            foreach ($line in ($result -split "`n")) { Log-Message ("[$Server] " + $line) }
        } catch {
            Log-Message "ERROR contacting ${Server} - $_"
        }
        Log-Message "Completed FlexLM license reset on ${Server}"
    }

    Log-Message '=== FLEXLM TEMP LICENSE RESET COMPLETED ==='
}

function Invoke-IISHoly3 {
    if (-not (Validate-Servers)) { return }
    $selectedServers = Get-AllSelectedServers
    $holy3 = @('ArchiveAppSvcAppPool','ArchiveViewsAppPool','ArchiveDataAppPool')

    Log-Message "=== IIS APP POOLS HOLY3 RECYCLE OPERATION STARTED ==="
    Log-Message "Selected Servers: $($selectedServers -join ', ')"
    Log-Message "App Pools (in order): $($holy3 -join ', ')"

    foreach ($Server in $selectedServers) {
        Log-Message "Processing HOLY3 on ${Server}..."
        foreach ($AppPool in $holy3) {
            try {
                Log-Message "Running: Restart-WebAppPool -Name $AppPool on $Server"
                $result = Invoke-Command -ComputerName $Server -ScriptBlock {
                    Param ($AppPool)
                    Import-Module WebAdministration
                    if (Test-Path "IIS:\\AppPools\\$AppPool") {
                        Restart-WebAppPool -Name $AppPool -ErrorAction Stop
                        return "[$env:COMPUTERNAME] ${AppPool} recycled successfully."
                    } else {
                        return "[$env:COMPUTERNAME] IIS App Pool ${AppPool} not found."
                    }
                } -ArgumentList $AppPool -ErrorAction Stop
                Log-Message $result
            } catch {
                Log-Message "ERROR recycling IIS App Pool ${AppPool} on ${Server} - $_"
            }
        }
        Log-Message "HOLY3 recycle processing completed for ${Server}"
    }
    Log-Message "=== IIS APP POOLS HOLY3 RECYCLE OPERATION COMPLETED ==="
}

# DICOM Node - Apply Changes (DAs): RESTART_SVC specific processes via AdminCommand and restart specific Windows service
function Invoke-DicomApplyChanges {
    if (-not (Validate-Servers)) { return }
    $selectedServers = Get-AllSelectedServers
    $processesToRestart = @('svdser','svdtc','svqe','svsm','svdsk','svfolder','svdds','svldr','svfwd')
    $windowsServiceToRestart = 'Imaginet Loader Server'

    Log-Message "=== DICOM APPLY CHANGES (DAs) STARTED ==="
    Log-Message "Selected Servers: $($selectedServers -join ', ')"
    Log-Message "Processes (RESTART_SVC): $($processesToRestart -join ', ')"
    Log-Message "Windows Service (RESTART): $windowsServiceToRestart"

    foreach ($Server in $selectedServers) {
        Log-Message "Restarting DICOM processes (AdminCommand RESTART_SERVICE) on ${Server}..."
        foreach ($Process in $processesToRestart) {
            try {
                $result = Invoke-Command -ComputerName $Server -ScriptBlock {
                    Param ($Process)
                    $imaginetRoot = [System.Environment]::GetEnvironmentVariable('imaginet_root')
                    $adminCmdPaths = if ($imaginetRoot) {
                        @(
                            (Join-Path $imaginetRoot 'utils\AdminCommand.exe'),
                            (Join-Path $imaginetRoot 'Portal\Dll\SVRender\AdminCommand.exe')
                        )
                    } else {
                        @('AdminCommand.exe')
                    }

                    foreach ($adminCmd in $adminCmdPaths) {
                        if ($adminCmd -eq 'AdminCommand.exe' -or (Test-Path $adminCmd)) {
                            try {
                                $output = & $adminCmd restart_service $Process 2>&1
                                $lines = @($output) | ForEach-Object { $_.ToString().Trim() } | Where-Object { $_ -ne '' }
                                $restarted = @($lines | Where-Object { $_ -match '^Restarted service' })
                                $notReq    = @($lines | Where-Object { $_ -match '^Restart not required' })
                                $failed    = @($lines | Where-Object { $_ -match '^Failed' -or $_ -match 'Error' })
                                $summary = "RESTART_SERVICE ${Process}: restarted=$($restarted.Count), not-required=$($notReq.Count)"
                                if ($failed.Count -gt 0) { $summary += "; FAILURES: " + ($failed -join ' | ') }
                                return "[$env:COMPUTERNAME] $summary"
                            } catch {
                                return "ERROR executing RESTART_SERVICE $Process via $adminCmd on $env:COMPUTERNAME - $_"
                            }
                        }
                    }
                    return "ERROR: AdminCommand.exe not found in any specified location on $env:COMPUTERNAME"
                } -ArgumentList $Process -ErrorAction Stop
                Log-Message $result
            } catch {
                Log-Message "ERROR executing RESTART_SERVICE for ${Process} on ${Server} - $_"
            }
        }

        # Restart the Windows service after processes
        try {
            $svcResult = Invoke-Command -ComputerName $Server -ScriptBlock {
                Param ($SvcName)
                $svc = Get-Service -Name $SvcName -ErrorAction SilentlyContinue
                if ($null -ne $svc) {
                    if ($svc.Status -eq 'Running') { Stop-Service -Name $SvcName -Force -ErrorAction SilentlyContinue }
                    Start-Sleep -Seconds 1
                    Start-Service -Name $SvcName -ErrorAction Stop
                    return "Service '$SvcName' restarted successfully on $env:COMPUTERNAME."
                } else { return "Service '$SvcName' not found on $env:COMPUTERNAME." }
            } -ArgumentList $windowsServiceToRestart -ErrorAction Stop
            Log-Message $svcResult
        } catch {
            Log-Message "ERROR restarting service '$windowsServiceToRestart' on ${Server} - $_"
        }

        Log-Message "DICOM Apply Changes processing completed for ${Server}"
    }
    Log-Message "=== DICOM APPLY CHANGES (DAs) COMPLETED ==="
}

# DICOM Node - Make Visible VMotion (Portals): recycle specific app pools
function Invoke-PortalMakeVisibleVMotion {
    if (-not (Validate-Servers)) { return }
    $selectedServers = Get-AllSelectedServers
    $appPools = @('ArchiveDataAppPool','ArchiveViewsAppPool','UIPatientsAppPool')

    Log-Message "=== PORTAL MAKE VISIBLE VMOTION STARTED ==="
    Log-Message "Selected Servers: $($selectedServers -join ', ')"
    Log-Message "App Pools: $($appPools -join ', ')"

    foreach ($Server in $selectedServers) {
        Log-Message "Recycling app pools on ${Server}..."
        foreach ($AppPool in $appPools) {
            try {
                $result = Invoke-Command -ComputerName $Server -ScriptBlock {
                    Param ($AppPool)
                    Import-Module WebAdministration
                    if (Test-Path "IIS:\AppPools\$AppPool") {
                        Restart-WebAppPool -Name $AppPool -ErrorAction Stop
                        return "[$env:COMPUTERNAME] $AppPool recycled successfully."
                    } else { return "[$env:COMPUTERNAME] IIS App Pool $AppPool not found." }
                } -ArgumentList $AppPool -ErrorAction Stop
                Log-Message $result
            } catch {
                Log-Message "ERROR recycling IIS App Pool ${AppPool} on ${Server} - $_"
            }
        }
        Log-Message "App pool recycle processing completed for ${Server}"
    }
    Log-Message "=== PORTAL MAKE VISIBLE VMOTION COMPLETED ==="
}

# AdminCommand status: show per-server parsed status in ListView windows
function Invoke-AdminStatus {
    if (-not (Validate-Servers)) { return }
    $selectedServers = Get-AllSelectedServers
    Log-Message "=== PROCESSES STATUS OPERATION STARTED ==="
    Log-Message "Selected Servers: $($selectedServers -join ', ')"

    # Progress UI
    $progressForm = New-Object System.Windows.Forms.Form
    $progressForm.Text = "Gathering Process Status..."
    $progressForm.Size = New-Object System.Drawing.Size(400, 150)
    $progressForm.StartPosition = "CenterParent"
    $progressForm.FormBorderStyle = "FixedDialog"
    $progressForm.MaximizeBox = $false
    $progressForm.MinimizeBox = $false
    $progressLabel = New-Object System.Windows.Forms.Label
    $progressLabel.Text = "Please wait while gathering process status from servers..."
    $progressLabel.Location = New-Object System.Drawing.Point(20, 30)
    $progressLabel.Size = New-Object System.Drawing.Size(350, 60)
    $progressLabel.TextAlign = "MiddleCenter"
    $progressForm.Controls.Add($progressLabel)
    [void]$progressForm.Show(); $progressForm.Refresh()

    $perServer = @{}
    foreach ($server in $selectedServers | Sort-Object) {
        try {
            $progressLabel.Text = "Gathering process status from: $server"; $progressForm.Refresh()
            Log-Message "Getting process status from $server..."
            $result = Invoke-Command -ComputerName $server -ScriptBlock {
                $imaginetRoot = [System.Environment]::GetEnvironmentVariable('imaginet_root')
                $adminCmdPaths = if ($imaginetRoot) {
                    @(
                        (Join-Path $imaginetRoot 'utils\AdminCommand.exe'),
                        (Join-Path $imaginetRoot 'Portal\Dll\SVRender\AdminCommand.exe')
                    )
                } else {
                    @('AdminCommand.exe')
                }

                $commandExecuted = $false
                $results = @()
                foreach ($adminCmd in $adminCmdPaths) {
                    try {
                        if ($adminCmd -eq 'AdminCommand.exe' -or (Test-Path $adminCmd)) {
                            $results += '=========================================='
                            $results += "SERVER: $env:COMPUTERNAME"
                            $results += "Command: $adminCmd status"
                            $results += '=========================================='
                            $out = & $adminCmd status 2>&1
                            $results += $out
                            $commandExecuted = $true
                            break
                        }
                    } catch {
                        $results += "Failed to execute $adminCmd status on $env:COMPUTERNAME - $_"
                    }
                }
                if (-not $commandExecuted) { return "ERROR: AdminCommand.exe not found in any specified location on $env:COMPUTERNAME" }
                return $results -join "`n"
            } -ErrorAction Stop
            $perServer[$server] = $result; Log-Message "Successfully retrieved process status from $server"
        } catch { $perServer[$server] = "ERROR connecting to $server - $_"; Log-Message "ERROR connecting to $server - $_" }
    }

    $progressForm.Close()

    foreach ($server in $selectedServers | Sort-Object) {
        $text = $perServer[$server]
        if (-not $text) { continue }

        # Parse AdminCommand status summary lines into rows
        $rows = @()
        $lines = $text -split "`n"
        foreach ($line in $lines) {
            $l = $line.Trim()
            if (-not $l) { continue }
            $m = [regex]::Match($l, '^\((?<svc>[^)]+)\):\s*Max Services:\s*(?<Max>\d+),\s*Pooled Services:\s*(?<Pooled>\d+),\s*Raised Services:\s*(?<Raised>\d+),\s*Idle Services:\s*(?<Idle>\d+),\s*Active Services:\s*(?<Active>\d+),\s*Pending Requests:\s*(?<PendingReq>\d+),\s*Total Requests:\s*(?<TotalReq>\d+)')
            if ($m.Success) {
                $rows += [PSCustomObject]@{
                    SERVICE    = $m.Groups['svc'].Value
                    Max        = [int]$m.Groups['Max'].Value
                    Pooled     = [int]$m.Groups['Pooled'].Value
                    Raised     = [int]$m.Groups['Raised'].Value
                    Idle       = [int]$m.Groups['Idle'].Value
                    Active     = [int]$m.Groups['Active'].Value
                    PendingReq = [int]$m.Groups['PendingReq'].Value
                    TotalReq   = [int]$m.Groups['TotalReq'].Value
                }
            }
        }

        # Per-server window
        $formSrv = New-Object System.Windows.Forms.Form
        $formSrv.Text = "Process Status - $server"
        $formSrv.Size = New-Object System.Drawing.Size(1100, 700)
        $formSrv.StartPosition = "CenterParent"
        $formSrv.FormBorderStyle = "Sizable"

        $lblHdr = New-Object System.Windows.Forms.Label
        $lblHdr.Text = "SERVICE status on $server"
        $lblHdr.Location = New-Object System.Drawing.Point(10, 10)
        $lblHdr.Size = New-Object System.Drawing.Size(800, 22)
        $lblHdr.Font = New-Object System.Drawing.Font("Arial", 11, [System.Drawing.FontStyle]::Bold)
        $formSrv.Controls.Add($lblHdr)

        $listView = New-Object System.Windows.Forms.ListView
        $listView.View = 'Details'
        $listView.FullRowSelect = $true
        $listView.GridLines = $true
        $listView.Location = New-Object System.Drawing.Point(10, 40)
        $listView.Size = New-Object System.Drawing.Size(1060, 560)
        $listView.Anchor = "Top, Left, Bottom, Right"
        $listView.Font = New-Object System.Drawing.Font("Consolas", 9)
        [void]$listView.Columns.Add("SERVICE", 260)
        [void]$listView.Columns.Add("Max", 60)
        [void]$listView.Columns.Add("Pooled", 70)
        [void]$listView.Columns.Add("Raised", 70)
        [void]$listView.Columns.Add("Idle", 60)
        [void]$listView.Columns.Add("Active", 70)
        [void]$listView.Columns.Add("PendingReq", 90)
        [void]$listView.Columns.Add("TotalReq", 100)

        foreach ($r in ($rows | Sort-Object SERVICE)) {
            $item = New-Object System.Windows.Forms.ListViewItem($r.SERVICE)
            [void]$item.SubItems.Add([string]$r.Max)
            [void]$item.SubItems.Add([string]$r.Pooled)
            [void]$item.SubItems.Add([string]$r.Raised)
            [void]$item.SubItems.Add([string]$r.Idle)
            [void]$item.SubItems.Add([string]$r.Active)
            [void]$item.SubItems.Add([string]$r.PendingReq)
            [void]$item.SubItems.Add([string]$r.TotalReq)
            [void]$listView.Items.Add($item)
        }

        $formSrv.Controls.Add($listView)
        $closeBtn = New-Object System.Windows.Forms.Button
        $closeBtn.Text = "Close"
        $closeBtn.Location = New-Object System.Drawing.Point(10, 610)
        $closeBtn.Size = New-Object System.Drawing.Size(80, 30)
        $closeBtn.Anchor = "Bottom, Left"
        $closeBtn.Add_Click({ $formSrv.Close() })
        $formSrv.Controls.Add($closeBtn)
        $formSrv.Show()
    }
    Log-Message "=== PROCESSES STATUS OPERATION COMPLETED ==="
}

# UPTIME REPORT: gather LastBootUpTime and uptime days per server and show in a simple table
function Invoke-UptimeReport {
    if (-not (Validate-Servers)) { return }
    $selectedServers = Get-AllSelectedServers
    Log-Message "=== UPTIME REPORT STARTED ==="
    Log-Message "Selected Servers: $($selectedServers -join ', ')"

    $results = @()
    foreach ($server in $selectedServers | Sort-Object) {
        try {
            Log-Message "Getting uptime from $server..."
            $data = Invoke-Command -ComputerName $server -ScriptBlock {
                $os = Get-CimInstance Win32_OperatingSystem
                $boot = $os.LastBootUpTime
                $bootStr = (Get-Date $boot -Format 'yyyy-MM-dd HH:mm:ss')
                $days = [math]::Round((New-TimeSpan -Start $boot -End (Get-Date)).TotalDays, 2)
                [pscustomobject]@{
                    Server = $env:COMPUTERNAME
                    LastBoot   = $bootStr
                    UptimeDays = $days
                }
            } -ErrorAction Stop
            if ($data) {
                $results += $data
                Log-Message "[$server] LastBoot: $($data.LastBoot) | UptimeDays: $($data.UptimeDays)"
            }
        } catch {
            Log-Message "ERROR getting uptime for ${server} - $_"
        }
    }

    # Results window
    $uf = New-Object System.Windows.Forms.Form
    $uf.Text = "Uptime Report"
    $uf.Size = New-Object System.Drawing.Size(700, 450)
    $uf.StartPosition = "CenterParent"
    $uf.FormBorderStyle = "Sizable"

    $lv = New-Object System.Windows.Forms.ListView
    $lv.View = 'Details'
    $lv.FullRowSelect = $true
    $lv.GridLines = $true
    $lv.Location = New-Object System.Drawing.Point(10, 10)
    $lv.Size = New-Object System.Drawing.Size(660, 360)
    $lv.Anchor = "Top, Left, Bottom, Right"
    [void]$lv.Columns.Add("Server", 200)
    [void]$lv.Columns.Add("Last Boot", 280)
    [void]$lv.Columns.Add("Uptime (days)", 140)

    foreach ($r in ($results | Sort-Object Server)) {
        $item = New-Object System.Windows.Forms.ListViewItem("$($r.Server)")
        [void]$item.SubItems.Add("$($r.LastBoot)")
        [void]$item.SubItems.Add("$($r.UptimeDays)")
        [void]$lv.Items.Add($item)
    }
    $uf.Controls.Add($lv)

    $btnClose = New-Object System.Windows.Forms.Button
    $btnClose.Text = "Close"
    $btnClose.Location = New-Object System.Drawing.Point(310, 380)
    $btnClose.Size = New-Object System.Drawing.Size(80, 25)
    $btnClose.Anchor = "Bottom"
    $btnClose.Add_Click({ $uf.Close() })
    $uf.Controls.Add($btnClose)

    [void]$uf.ShowDialog()

    Log-Message "=== UPTIME REPORT COMPLETED ==="
}

# CSVER REPORT function (migrated from button handler)
function Invoke-CsverReport {
    if (-not (Validate-Servers)) { return }
    $selectedServers = Get-AllSelectedServers
    Log-Message '=== CSVER REPORT OPERATION STARTED ==='
    Log-Message "Selected Servers: $($selectedServers -join ', ')"

    # Progress form
    $progressForm = New-Object System.Windows.Forms.Form
    $progressForm.Text = "Gathering CSVER Information..."
    $progressForm.Size = New-Object System.Drawing.Size(400, 150)
    $progressForm.StartPosition = "CenterParent"
    $progressForm.FormBorderStyle = "FixedDialog"
    $progressForm.MaximizeBox = $false
    $progressForm.MinimizeBox = $false
    $progressLabel = New-Object System.Windows.Forms.Label
    $progressLabel.Text = "Please wait while gathering version information from servers..."
    $progressLabel.Location = New-Object System.Drawing.Point(20, 30)
    $progressLabel.Size = New-Object System.Drawing.Size(350, 60)
    $progressLabel.TextAlign = "MiddleCenter"
    $progressForm.Controls.Add($progressLabel)
    [void]$progressForm.Show(); $progressForm.Refresh()

    $serverData = @()
    $allVersions = @{
        'Vue PACS (WFM) Server' = @{}
        'Vue PACS (WFM) Cluster - Primary Node' = @{}
        'Vue PACS (WFM) Cluster - Additional Node' = @{}
        'Vue PACS (WFM) Application Server' = @{}
        'Vue PACS Client' = @{}
        'NDF' = @{}
        'Vue Motion' = @{}
        'Vue Explorer' = @{}
        'Web Reporting' = @{}
        'Diagnostic Viewer Settings' = @{}
        'Web System Configuration' = @{}
    }

    foreach ($server in $selectedServers) {
        try {
            $progressLabel.Text = "Gathering CSVER info from: $server"
            $progressForm.Refresh()
            Log-Message "Getting CSVER information from $server..."
            $result = Invoke-Command -ComputerName $server -ScriptBlock {
                try {
                    $ErrorActionPreference = 'Continue'
                    # Resolve python/py if available
                    $py = $null
                    $pc = Get-Command python -ErrorAction SilentlyContinue; if ($pc) { $py = $pc.Source }
                    if (-not $py) { $pc = Get-Command py -ErrorAction SilentlyContinue; if ($pc) { $py = $pc.Source } }

                    # Prefer csver.py in known paths
                    $csverPath = $null
                    if (Test-Path 'C:\\Program Files\\Carestream\\System5\\scripts\\csver.py') { $csverPath = 'C:\\Program Files\\Carestream\\System5\\scripts\\csver.py' }
                    elseif (Test-Path 'C:\\Program Files\\Philips\\System5\\scripts\\csver.py') { $csverPath = 'C:\\Program Files\\Philips\\System5\\scripts\\csver.py' }

                    $output = $null
                    if ($csverPath) {
                        if ($py) { $output = & $py $csverPath 2>&1 }
                        else     { $output = & $csverPath 2>&1 }
                    }
                    else {
                        $cmd = $null
                        if (Get-Command csver.py -ErrorAction SilentlyContinue) { $cmd = 'csver.py' }
                        elseif (Get-Command csver -ErrorAction SilentlyContinue) { $cmd = 'csver' }
                        if ($cmd) { $output = & $cmd 2>&1 }
                        else { $output = 'csver not found on PATH or known locations' }
                    }

                    return @{ Success = $true; Output = $output; Hostname = ${env:COMPUTERNAME} }
                } catch {
                    return @{ Success = $false; Error = $_.Exception.Message; Hostname = ${env:COMPUTERNAME} }
                }
            } -ErrorAction Stop

            # Resolve properties from possible hashtable/PSObject and ensure hostname fallback
            $resolvedHostname = $null
            try { if ($null -ne $result.Hostname -and $result.Hostname -ne '') { $resolvedHostname = $result.Hostname } } catch {}
            if ([string]::IsNullOrWhiteSpace($resolvedHostname)) { $resolvedHostname = $server }

            $isSuccess = $false
            try { if ($result.Success) { $isSuccess = $true } } catch {}

            if ($isSuccess) {
                $serverInfo = [pscustomobject]@{ Hostname = $resolvedHostname; RequestedName = $server; Versions = @{}; Patches = @(); RawOutput = ($result.Output -join "`n"); Error = $null }
                $inPatchSection = $false
                # Normalize remote output into string lines robustly
                $__lines = @()
                if ($null -eq $result.Output) {
                    $__lines = @()
                } elseif ($result.Output -is [string]) {
                    $__lines = $result.Output -split "`r?`n"
                } elseif ($result.Output -is [System.Collections.IEnumerable]) {
                    # If it's already an array of strings, Out-String will join them safely; if it includes objects, ToString will render them
                    $__text = ($result.Output | Out-String)
                    $__lines = $__text -split "`r?`n"
                } else {
                    $__lines = ([string]$result.Output) -split "`r?`n"
                }
                # Build a single, shared component line regex (case-insensitive) and allow optional 'released:' segment
                $componentPattern = "(?i)^(Vue PACS \(WFM\) Server|Vue PACS \(WFM\) Cluster - Primary Node|Vue PACS \(WFM\) Cluster - Additional Node|Vue PACS \(WFM\) Application Server|Vue PACS Client|NDF|Vue Motion|Vue Explorer|Web Reporting|Diagnostic Viewer Settings|Web System Configuration)\s+v?([\d.]+)\s+\(build\s*([0-9]+)(?:,\s*released\s*:?\s*(.+))?\)"
                foreach ($line in $__lines) {
                    $line = ([string]$line).Trim()
                    # Detect start of patches section(s), case-insensitive
                    if ($line -match "(?i)^(server\s+patches|portal\s+patches):") { $inPatchSection = $true; continue }
                    # If we reach a component line or blank line, end patches section
                    if ($inPatchSection -and (($line -eq '') -or ($line -match $componentPattern))) { $inPatchSection = $false; if ($line -eq '') { continue } }
                    if ($inPatchSection) {
                        if (-not [string]::IsNullOrWhiteSpace($line)) { $serverInfo.Patches += $line }
                        continue
                    }
                    if ($line -match $componentPattern) {
                        $component,$version,$build,$releaseDate = $matches[1],$matches[2],$matches[3],$matches[4]
                        # If release date was not present, omit it from the version string
                        if ([string]::IsNullOrWhiteSpace($releaseDate)) {
                            $versionInfo = "$version (build $build)"
                        } else {
                            $versionInfo = "$version (build $build, released: $releaseDate)"
                        }
                        $serverInfo.Versions[$component] = $versionInfo
                        if (-not $allVersions[$component].ContainsKey($versionInfo)) { $allVersions[$component][$versionInfo] = @() }
                        $allVersions[$component][$versionInfo] += $resolvedHostname
                    }
                }
                $serverData += $serverInfo
                try { Log-Message ("[CSVER] {0} patches collected for {1}" -f $serverInfo.Patches.Count, $resolvedHostname) } catch {}
                Log-Message "Successfully retrieved CSVER info from $server"
            } else {
                $errMsg = $null; try { $errMsg = $result.Error } catch { $errMsg = $_ }
                $serverData += ([pscustomobject]@{ Hostname = $resolvedHostname; RequestedName = $server; Versions = @{}; Patches = @(); RawOutput = ""; Error = $errMsg })
                Log-Message "ERROR getting CSVER info from $server - $($result.Error)"
            }
        } catch {
            $serverData += ([pscustomobject]@{ Hostname = $server; RequestedName = $server; Versions = @{}; Patches = @(); RawOutput = ""; Error = "Failed to connect or execute command - $_" })
            Log-Message "ERROR connecting to $server - $_"
        }
    }

    $progressForm.Close()

    # Build CSVER report text
    $report = @()
    $report += "=== CSVER REPORT ==="
    $report += "Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    $report += "Servers: $($selectedServers.Count) selected"
    $report += ""

    # Summary table (servers across the top, rows per component-version)
    # Desired component order: WFM at top, then Vue Motion, Vue Explorer, Client, Diagnostic Viewer, Web Reporting, others
    $orderedComponents = @(
        'Vue PACS (WFM) Server',
        'Vue PACS (WFM) Cluster - Primary Node',
        'Vue PACS (WFM) Cluster - Additional Node',
        'Vue PACS (WFM) Application Server',
        'Vue Motion',
        'Vue Explorer',
        'Vue PACS Client',
        'Diagnostic Viewer Settings',
        'Web Reporting',
        'Web System Configuration',
        'NDF'
    )

    # Any additional components discovered that aren't in the preferred list
    $existingComponents = $allVersions.Keys
    $additionalComponents = @()
    foreach ($c in ($existingComponents | Sort-Object)) {
        if ($orderedComponents -notcontains $c) { $additionalComponents += $c }
    }
    $finalComponents = @()
    foreach ($c in $orderedComponents) { if ($existingComponents -contains $c) { $finalComponents += $c } }
    $finalComponents += $additionalComponents

    # Map server -> data for quick access and determine server column order
    $serverMap = @{}
    foreach ($sd in $serverData) { $serverMap[$sd.Hostname] = $sd }
    $serversSorted = ($serverData | ForEach-Object { $_.Hostname } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Sort-Object -Unique)
    if ($serversSorted.Count -eq 0) {
        $fallback = @()
        foreach ($comp in $allVersions.Keys) {
            foreach ($ver in $allVersions[$comp].Keys) { $fallback += $allVersions[$comp][$ver] }
        }
        $serversSorted = ($fallback | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Sort-Object -Unique)
    }

    # Reorder servers by group: CORE, DA, PORTAL, OTHER
    $core = @(); $da = @(); $portal = @(); $other = @(); $unknown = @()
    foreach ($sv in $serversSorted) {
        if ($global:CurrentCoreServers -and ($global:CurrentCoreServers -contains $sv)) { $core += $sv; continue }
        if ($global:CurrentDAServers -and ($global:CurrentDAServers -contains $sv)) { $da += $sv; continue }
        if ($global:CurrentPortalServers -and ($global:CurrentPortalServers -contains $sv)) { $portal += $sv; continue }
        if ($global:CurrentOtherServers -and ($global:CurrentOtherServers -contains $sv)) { $other += $sv; continue }
        $unknown += $sv
    }
    $serversSorted = @() + $core + $da + $portal + $other + $unknown

    # Helper to CSV-escape values
    function _CsvEscape([string]$val) {
        if ($null -eq $val) { $val = '' }
        $escaped = $val -replace '"', '""'
        return '"' + $escaped + '"'
    }

    # Helper to trim release date from version string for compact cells
    function _TrimVersion([string]$ver) {
        if ([string]::IsNullOrWhiteSpace($ver)) { return '' }
        $ver = $ver.Trim()
        # Extract build number if present
        $build = $null
        if ($ver -match '\(build\s*([0-9]+)') { $build = $matches[1] }
        # Extract the version text before '(build'
        $verOnly = $ver
        $idx = $ver.IndexOf('(')
        if ($idx -gt 0) { $verOnly = $ver.Substring(0, $idx).Trim() }
        # Remove leading 'v'
        if ($verOnly -like 'v*') { $verOnly = $verOnly.Substring(1) }
        if ([string]::IsNullOrWhiteSpace($verOnly)) { $verOnly = '' }
        if ([string]::IsNullOrWhiteSpace($build)) {
            return $verOnly
        } else {
            return "$verOnly (build $build)"
        }
    }

    $tableCsv = @()
    # Header row: Component, Version (build), then per-server checkbox columns
    $header = @('Component','Version (build)')
    foreach ($sv in $serversSorted) { $header += $sv }
    $tableCsv += ($header | ForEach-Object { _CsvEscape $_ }) -join ','

    foreach ($comp in $finalComponents) {
        if (-not $allVersions.ContainsKey($comp)) { continue }
        $versionsMap = $allVersions[$comp]
        if ($null -eq $versionsMap -or $versionsMap.Count -eq 0) { continue }
        foreach ($verKey in ($versionsMap.Keys | Sort-Object)) {
            $row = @($comp, (_TrimVersion $verKey))
            foreach ($sv in $serversSorted) {
                $has = ''
                if ($versionsMap[$verKey] -contains $sv) { $has = 'Y' }
                $row += $has
            }
            $tableCsv += ($row | ForEach-Object { _CsvEscape $_ }) -join ','
        }
    }

    # Single patches row with full list from each server
    $patchRow = @('Patches','')
    foreach ($sv in $serversSorted) {
        $patches = @()
        if ($serverMap.ContainsKey($sv)) { $patches = $serverMap[$sv].Patches }
        $patchCell = ($patches -join "`n")
        $patchRow += $patchCell
    }
    $tableCsv += ($patchRow | ForEach-Object { _CsvEscape $_ }) -join ','

    $report += '=== SUMMARY (TABLE) ==='
    $report += '(CSV format; first column is Component, second is Version (build), then one column per server; server column shows Y if that server has that version)'
    $report += ''
    $report += $tableCsv
    $report += ''

    # Optional: keep the original summary by component for reference
    $report += '=== SUMMARY BY COMPONENT ==='
    foreach ($component in ($allVersions.Keys | Sort-Object)) {
        $report += "--- $component ---"
        $versions = $allVersions[$component]
        if ($versions.Count -eq 0) { $report += "  No data"; $report += ""; continue }
        foreach ($ver in ($versions.Keys | Sort-Object)) {
            $svrs = $versions[$ver] | Sort-Object
            $report += "  $ver"
            $report += "    Servers: " + ($svrs -join ', ')
        }
        $report += ""
    }

    # Detailed per-server
    $report += "=== DETAILED RESULTS BY SERVER ==="
    $report += ""
    foreach ($s in ($serverData | Sort-Object Hostname)) {
        $nameToShow = if ($s.PSObject.Properties['Hostname'] -and -not [string]::IsNullOrWhiteSpace($s.Hostname)) { $s.Hostname } elseif ($s.PSObject.Properties['RequestedName'] -and -not [string]::IsNullOrWhiteSpace($s.RequestedName)) { $s.RequestedName } else { '(unknown)' }
        $report += "===== Results for ${nameToShow} ====="
        if ($s.Error) { $report += "ERROR: $($s.Error)"; $report += ""; continue }
        if ($s.Versions.Keys.Count -gt 0) {
            foreach ($k in ($s.Versions.Keys | Sort-Object)) {
                $report += ("{0}" -f $k).PadRight(45) + $s.Versions[$k]
            }
        } else {
            $report += "No component versions detected"
            # Show raw captured output to aid debugging when parsing finds nothing (limit to 200 lines)
            if ($s.RawOutput -and -not [string]::IsNullOrWhiteSpace($s.RawOutput)) {
                $report += "(Raw output from server follows)"
                $rawLines = [System.Environment]::NewLine
                try { $rawLines = $s.RawOutput -split "`r?`n" } catch { $rawLines = @($s.RawOutput) }
                $take = [Math]::Min(200, $rawLines.Count)
                foreach ($rl in $rawLines[0..($take-1)]) { $report += "  $rl" }
                if ($rawLines.Count -gt $take) { $report += "  ... ($($rawLines.Count - $take) more lines truncated)" }
            }
        }
        if ($s.Patches.Count -gt 0) {
            $report += "Patches:"
            $report += ($s.Patches | ForEach-Object { "  $_" })
        }
        $report += ""
    }

    # Build DataTable for grid view (rows per component-version), server checkbox columns
    $dt = New-Object System.Data.DataTable
    [void]$dt.Columns.Add('Component', [string])
    [void]$dt.Columns.Add('Version (build)', [string])
    foreach ($sv in $serversSorted) {
        if ([string]::IsNullOrWhiteSpace($sv)) { continue }
        $colHas = New-Object System.Data.DataColumn($sv, [bool])
        $colHas.AllowDBNull = $true
        [void]$dt.Columns.Add($colHas)
    }

    foreach ($comp in $finalComponents) {
        $versionsMap = $allVersions[$comp]
        if ($versionsMap.Count -eq 0) { continue }
        foreach ($verKey in ($versionsMap.Keys | Sort-Object)) {
            $r = $dt.NewRow()
            $r['Component'] = $comp
            $r['Version (build)'] = _TrimVersion $verKey
            foreach ($sv in $serversSorted) {
                if (-not [string]::IsNullOrWhiteSpace($sv)) {
                    $r[$sv] = [bool]($versionsMap[$verKey] -contains $sv)
                }
            }
            [void]$dt.Rows.Add($r)
        }
    }
    # Add patches as a final row in main grid (newline-separated per server)
    $rP = $dt.NewRow()
    $rP['Component'] = 'Patches'
    $rP['Version (build)'] = ''
    foreach ($sv in $serversSorted) {
        # store as DBNull for checkbox data type; we'll paint text via CellPainting
        $rP[$sv] = [DBNull]::Value
    }
    [void]$dt.Rows.Add($rP)

    # Results window
    $resultsForm = New-Object System.Windows.Forms.Form
    $resultsForm.Text = "CSVER Report (Enhanced) - $($selectedServers.Count) Servers"
    $resultsForm.Size = New-Object System.Drawing.Size(1400, 700)
    $resultsForm.StartPosition = "CenterParent"
    $resultsForm.FormBorderStyle = "Sizable"

    # Tabs: Summary Table and Full Text Report
    $tabControl = New-Object System.Windows.Forms.TabControl
    $tabControl.Location = New-Object System.Drawing.Point(10, 10)
    $tabControl.Size = New-Object System.Drawing.Size(1360, 600)
    $tabControl.Anchor = "Top, Left, Bottom, Right"

    $tabSummary = New-Object System.Windows.Forms.TabPage
    $tabSummary.Text = 'Summary Table'
    # Visual marker label to confirm enhanced UI path is in effect
    $markerLbl = New-Object System.Windows.Forms.Label
    $markerLbl.AutoSize = $true
    $markerLbl.Text = 'Enhanced Summary Grid Active'
    $markerLbl.ForeColor = [System.Drawing.Color]::DarkGreen
    $markerLbl.Font = New-Object System.Drawing.Font('Arial', 8, [System.Drawing.FontStyle]::Bold)
    $markerLbl.Dock = 'Top'
    $markerLbl.Padding = '6,3,6,3'
    # Legend for server group colors
    $legendPanel = New-Object System.Windows.Forms.Panel
    $legendPanel.Dock = 'Top'
    $legendPanel.Height = 26
    $legendPanel.Padding = '6,3,6,3'
    function New-GroupLabel([string]$text, [System.Drawing.Color]$color) {
        $lbl = New-Object System.Windows.Forms.Label
        $lbl.Text = $text
        $lbl.AutoSize = $true
        $lbl.Margin = '8,0,12,0'
        $lbl.Padding = '6,2,6,2'
        $lbl.BackColor = $color
        $lbl.BorderStyle = 'FixedSingle'
        return $lbl
    }
    $legendPanel.Controls.Add((New-GroupLabel 'CORE' ([System.Drawing.Color]::FromArgb(230,245,255))))
    $legendPanel.Controls.Add((New-GroupLabel 'DA' ([System.Drawing.Color]::FromArgb(240,255,230))))
    $legendPanel.Controls.Add((New-GroupLabel 'PORTAL' ([System.Drawing.Color]::FromArgb(255,240,230))))
    $legendPanel.Controls.Add((New-GroupLabel 'OTHER' ([System.Drawing.Color]::FromArgb(245,240,255))))
    # Group count banner for quick verification of group matching
    $groupCount = New-Object System.Windows.Forms.Label
    $groupCount.Dock = 'Top'
    $groupCount.AutoSize = $true
    $groupCount.Padding = '6,2,6,2'
    $countCore = ($serversSorted | Where-Object { $global:CurrentCoreServers -and ($global:CurrentCoreServers -contains $_) }).Count
    $countDA = ($serversSorted | Where-Object { $global:CurrentDAServers -and ($global:CurrentDAServers -contains $_) }).Count
    $countPortal = ($serversSorted | Where-Object { $global:CurrentPortalServers -and ($global:CurrentPortalServers -contains $_) }).Count
    $countOther = ($serversSorted | Where-Object { $global:CurrentOtherServers -and ($global:CurrentOtherServers -contains $_) }).Count
    $groupCount.Text = "Groups: CORE=$countCore | DA=$countDA | PORTAL=$countPortal | OTHER=$countOther"
    # Container for top markers to avoid overlap with grid
    $headerPanel = New-Object System.Windows.Forms.FlowLayoutPanel
    $headerPanel.Dock = 'Top'
    $headerPanel.AutoSize = $true
    $headerPanel.AutoSizeMode = 'GrowAndShrink'
    $headerPanel.WrapContents = $true
    $headerPanel.Padding = '6,3,6,3'
    $headerPanel.FlowDirection = 'LeftToRight'
    [void]$headerPanel.Controls.Add($markerLbl)
    [void]$headerPanel.Controls.Add($groupCount)
    $grid = New-Object System.Windows.Forms.DataGridView
    $grid.Dock = 'Fill'
    $grid.ReadOnly = $true
    $grid.AllowUserToAddRows = $false
    $grid.AllowUserToDeleteRows = $false
    $grid.AutoSizeColumnsMode = 'None'
    $grid.EnableHeadersVisualStyles = $false
    $grid.AllowUserToOrderColumns = $false
    $grid.RowHeadersVisible = $false
    $grid.ColumnHeadersVisible = $true
    $grid.AutoGenerateColumns = $true
    $grid.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::Gainsboro
    $grid.ColumnHeadersDefaultCellStyle.ForeColor = [System.Drawing.Color]::Black
    $grid.ColumnHeadersDefaultCellStyle.Alignment = [System.Windows.Forms.DataGridViewContentAlignment]::MiddleCenter
    $grid.ColumnHeadersDefaultCellStyle.Font = (New-Object System.Drawing.Font($grid.Font, [System.Drawing.FontStyle]::Bold))
    $grid.ColumnHeadersHeightSizeMode = 'EnableResizing'
    $grid.ColumnHeadersHeight = 28
    $grid.ColumnHeadersBorderStyle = 'Single'
    # Default: do not wrap cell text; patches row will also be non-wrapping
    $grid.DefaultCellStyle.WrapMode = 'False'
    # Allow manual row resizing (especially Patches row)
    $grid.AutoSizeRowsMode = 'None'
    $grid.AllowUserToResizeRows = $true
    $grid.AllowUserToResizeColumns = $true
    # Set desired initial widths (user may edit these two lines to change defaults)
    $InitialComponentWidth = 220
    $InitialVersionWidth = 120
    # Exposed config: initial Patches row height (~multiplier x normal row height)
    $PatchesInitialHeightMultiplier = 8
    $PatchesInitialHeightMinPx = 80
    # Pass server context and CSV lines to event handlers via Tag
    $grid.Tag = [pscustomobject]@{ ServerOrder = $serversSorted; ServerMap = $serverMap; CsvLines = $tableCsv; DoneSetup = $false; WidthSync = $false; PatchesHeightMult = $PatchesInitialHeightMultiplier; PatchesMinPx = $PatchesInitialHeightMinPx }

    # Helper to apply post-bind setup reliably (can be called from multiple events and after binding)
    $applyGridSetup = {
        param($g)
        try { Log-Message "[CSVER] applyGridSetup invoked" } catch {}
        if ($null -eq $g -or $g.Columns.Count -eq 0) { return }
        if ($g.Tag -and $g.Tag.DoneSetup) { return }
        $svOrder = $g.Tag.ServerOrder
        # Convert server columns to checkbox type and color by group
        foreach ($sv in $svOrder) {
            $col = $g.Columns[$sv]
            if ($null -ne $col) {
                if ($col.ValueType -ne [bool]) {
                    $idx = $col.Index
                    $cb = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
                    $cb.Name = $col.Name
                    $cb.HeaderText = $col.HeaderText
                    $cb.DataPropertyName = $col.DataPropertyName
                    $cb.ThreeState = $true
                    $cb.DefaultCellStyle.NullValue = $null
                    $g.Columns.RemoveAt($idx)
                    $g.Columns.Insert($idx, $cb)
                    $col = $g.Columns[$idx]
                }
                $h = $col.HeaderCell
                $h.Style.ForeColor = [System.Drawing.Color]::Black
                $h.Style.Alignment = [System.Windows.Forms.DataGridViewContentAlignment]::MiddleCenter
                $h.Style.Font = (New-Object System.Drawing.Font($g.Font, [System.Drawing.FontStyle]::Bold))
                if ($global:CurrentCoreServers -and ($global:CurrentCoreServers -contains $sv)) { $h.Style.BackColor = [System.Drawing.Color]::FromArgb(230,245,255); $col.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(245,250,255) }
                elseif ($global:CurrentDAServers -and ($global:CurrentDAServers -contains $sv)) { $h.Style.BackColor = [System.Drawing.Color]::FromArgb(240,255,230); $col.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(248,255,240) }
                elseif ($global:CurrentPortalServers -and ($global:CurrentPortalServers -contains $sv)) { $h.Style.BackColor = [System.Drawing.Color]::FromArgb(255,240,230); $col.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(255,248,240) }
                elseif ($global:CurrentOtherServers -and ($global:CurrentOtherServers -contains $sv)) { $h.Style.BackColor = [System.Drawing.Color]::FromArgb(245,240,255); $col.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(248,246,255) }
            }
        }
        # Make Version (build) and Component columns wider and resizable; disable sorting on all columns
        $vb = $g.Columns['Version (build)']
        if ($vb) {
            $vb.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::None
            $vb.Width = $InitialVersionWidth
            $vb.SortMode = [System.Windows.Forms.DataGridViewColumnSortMode]::NotSortable
            $vb.Frozen = $true
        }
        $compCol = $g.Columns['Component']
        if ($compCol) {
            $compCol.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::None
            $compCol.Width = $InitialComponentWidth
            $compCol.SortMode = [System.Windows.Forms.DataGridViewColumnSortMode]::NotSortable
            $compCol.Frozen = $true
            $compCol.HeaderCell.Style.Alignment = [System.Windows.Forms.DataGridViewContentAlignment]::MiddleCenter
            $compCol.HeaderCell.Style.Font = (New-Object System.Drawing.Font($g.Font, [System.Drawing.FontStyle]::Bold))
        }
        foreach ($c in $g.Columns) { if ($c) { $c.SortMode = [System.Windows.Forms.DataGridViewColumnSortMode]::NotSortable } }
        # Replace last row (Patches) server cells with TextBoxCells containing multi-line text
        if ($g.Rows.Count -gt 0) {
            $pi = $g.Rows.Count - 1
            foreach ($sv in $svOrder) {
                $col = $g.Columns[$sv]
                if ($null -ne $col) {
                    $text = ''
                    if ($g.Tag.ServerMap.ContainsKey($sv)) { $text = ($g.Tag.ServerMap[$sv].Patches -join "`n") }
                    $tb = New-Object System.Windows.Forms.DataGridViewTextBoxCell
                    $tb.Value = $text
                    $g.Rows[$pi].Cells[$col.Index] = $tb
                    $g.Rows[$pi].Cells[$col.Index].Style.WrapMode = 'False'
                    $g.Rows[$pi].Cells[$col.Index].Style.Alignment = [System.Windows.Forms.DataGridViewContentAlignment]::TopLeft
                }
            }
        }
        # Freeze the top data row as requested
        if ($g.Rows.Count -gt 0) { $g.Rows[0].Frozen = $true }
        # Top-justify and size the patches row
        if ($g.Rows.Count -gt 0) {
            $pi = $g.Rows.Count - 1
            $g.Rows[$pi].DefaultCellStyle.Alignment = [System.Windows.Forms.DataGridViewContentAlignment]::TopLeft
            $g.Rows[$pi].DefaultCellStyle.WrapMode = 'False'
            $g.Rows[$pi].DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(255, 253, 240)
            # Initial height using exposed multiplier and minimum
            $lineHeight = $g.Font.Height + 6
            $mult = if ($g.Tag.PatchesHeightMult) { [double]$g.Tag.PatchesHeightMult } else { 8 }
            $minPx = if ($g.Tag.PatchesMinPx) { [int]$g.Tag.PatchesMinPx } else { 80 }
            $minHeight = [Math]::Max([int]([Math]::Ceiling($mult * $lineHeight)), $minPx)
            $g.Rows[$pi].Resizable = [System.Windows.Forms.DataGridViewTriState]::True
            $g.Rows[$pi].MinimumHeight = $minHeight
            $g.Rows[$pi].Height = [Math]::Min(800, [Math]::Max($g.Rows[$pi].Height, $minHeight))
        }
        if ($g.Tag) { $g.Tag.DoneSetup = $true }
        # Reassert patches row height after layout completes
    $null = $g.BeginInvoke([System.Action]{ try { if ($g.Rows.Count -gt 0) { $pi=$g.Rows.Count-1; $lh=$g.Font.Height+6; $mult=([double]($g.Tag.PatchesHeightMult)); if ($mult -le 0) { $mult=8 } $minPx=([int]($g.Tag.PatchesMinPx)); if ($minPx -le 0) { $minPx=80 } $mh=[Math]::Max([int]([Math]::Ceiling($mult*$lh)),$minPx); $g.Rows[$pi].MinimumHeight=$mh; if ($g.Rows[$pi].Height -lt $mh) { $g.Rows[$pi].Height=$mh } } } catch {} })
        try { Log-Message "[CSVER] applyGridSetup completed" } catch {}
    }
    # Color headers by group and set checkbox columns
    $grid.Add_DataBindingComplete({ & $applyGridSetup $grid })
    $grid.Add_DataSourceChanged({ & $applyGridSetup $grid })
    # Sync server column widths when any one server column is resized
    $grid.Add_ColumnWidthChanged({
        param($s, $e)
        try {
            $svOrder = $grid.Tag.ServerOrder
            if (-not $svOrder) { return }
            $col = $e.Column
            if (-not $col) { return }
            if ($grid.Tag.WidthSync) { return }
            if ($col.Name -and ($svOrder -contains $col.Name)) {
                $grid.Tag.WidthSync = $true
                $w = $col.Width
                foreach ($sv in $svOrder) {
                    $c2 = $grid.Columns[$sv]
                    if ($c2 -and $c2.Width -ne $w) { $c2.Width = $w }
                }
            }
        } finally { $grid.Tag.WidthSync = $false }
    })
    # Suppress DataGridView data errors from type formatting mismatches (attach early)
    $grid.Add_DataError({
        param($s,$a)
        try { Log-Message ("[CSVER] DataGridView DataError (prebind): {0}" -f $a.Exception.Message) } catch {}
        $a.ThrowException = $false
    })

    # Ensure patches row shows text even if column cell types revert; draw a single-line, clipped string
    $grid.Add_CellPainting({
        param($src, $cea)
        if ($cea.RowIndex -lt 0 -or $cea.ColumnIndex -lt 0) { return }
        $isPatchesRow = ($cea.RowIndex -eq ($grid.Rows.Count - 1))
        if (-not $isPatchesRow) { return }
        $colName = $grid.Columns[$cea.ColumnIndex].Name
        $svOrder = $grid.Tag.ServerOrder
        if ($svOrder -notcontains $colName) { return }

        # Background
        $cea.PaintBackground($cea.CellBounds, $true) | Out-Null
        # Draw each patch on its own line, no wrapping, clipped if too long
        $patches = @()
        if ($grid.Tag.ServerMap.ContainsKey($colName)) { $patches = $grid.Tag.ServerMap[$colName].Patches }
        $rect = $cea.CellBounds
        $padding = 4
        $x = $rect.X + $padding
        $y = $rect.Y + $padding
        $width = [Math]::Max(0, $rect.Width - (2*$padding))
        $height = [Math]::Max(0, $rect.Height - (2*$padding))
        $lineH = $grid.Font.Height + 2
        $flags = [System.Windows.Forms.TextFormatFlags]::SingleLine -bor [System.Windows.Forms.TextFormatFlags]::EndEllipsis -bor [System.Windows.Forms.TextFormatFlags]::Left -bor [System.Windows.Forms.TextFormatFlags]::Top -bor [System.Windows.Forms.TextFormatFlags]::NoPrefix
        foreach ($p in $patches) {
            if (($y - $rect.Y) -ge $height) { break }
            $lineRect = New-Object System.Drawing.Rectangle($x, $y, $width, $lineH)
            [System.Windows.Forms.TextRenderer]::DrawText($cea.Graphics, [string]$p, $grid.Font, $lineRect, $grid.ForeColor, $flags)
            $y += $lineH
        }
        $cea.Handled = $true
    })

    # Bind data AFTER wiring the handlers; also explicitly invoke setup after binding as a fallback
    $grid.DataSource = $dt
    try { Log-Message "[CSVER] DataSource assigned to summary grid" } catch {}
    # Run setup immediately in case events are suppressed
    & $applyGridSetup $grid
    # Post to message loop to ensure columns are created before setup runs
    $null = $resultsForm.BeginInvoke([System.Action]{ & $applyGridSetup $grid })
    # One-shot delayed setup to defeat any downstream overrides
    $once = New-Object System.Windows.Forms.Timer
    $once.Interval = 200
    $once.Add_Tick({ try { Log-Message "[CSVER] one-shot timer applying grid setup" } catch {}; & $applyGridSetup $grid; $once.Stop(); $once.Dispose() })
    $once.Start()
    # Reinforce header visibility post-bind
    $grid.ColumnHeadersVisible = $true
    $grid.RowHeadersVisible = $false
    $grid.ColumnHeadersHeight = [Math]::Max(24, $grid.ColumnHeadersHeight)
    $grid.Invalidate(); $grid.Refresh()

    # No custom paint needed now; patches are written into text cells in the last row

    # Add top header panel and legend (Top), then grid (Fill) so layout doesn't overlap
    $tabSummary.Controls.Add($grid)
    # Legend removed per user request
    # $tabSummary.Controls.Add($legendPanel)
    $tabSummary.Controls.Add($headerPanel)

    $tabText = New-Object System.Windows.Forms.TabPage
    $tabText.Text = 'Full Text Report'
    $resultsTextBox = New-Object System.Windows.Forms.TextBox
    $resultsTextBox.Multiline = $true
    $resultsTextBox.ScrollBars = 'Vertical'
    $resultsTextBox.Font = New-Object System.Drawing.Font('Consolas', 9)
    $resultsTextBox.WordWrap = $false
    $resultsTextBox.ReadOnly = $true
    # Only include the detailed section in Full Text Report
    $detailOnly = $report
    $idxHdr = ($report | Select-String -SimpleMatch '=== DETAILED RESULTS BY SERVER ===' | Select-Object -First 1).LineNumber
    if ($idxHdr) { $detailOnly = $report[($idxHdr-1)..($report.Count-1)] }
    $resultsTextBox.Text = $detailOnly -join "`r`n"
    $resultsTextBox.SelectionStart = 0; $resultsTextBox.SelectionLength = 0
    $resultsTextBox.Dock = 'Fill'
    $tabText.Controls.Add($resultsTextBox)

    [void]$tabControl.TabPages.Add($tabSummary)
    [void]$tabControl.TabPages.Add($tabText)
    # Re-apply styling when switching back to Summary
    $tabControl.Add_SelectedIndexChanged({ if ($tabControl.SelectedTab -eq $tabSummary) { try { Log-Message '[CSVER] Re-applying setup on Summary tab select' } catch {}; & $applyGridSetup $grid; $grid.Invalidate(); $grid.Refresh() } })
    $resultsForm.Controls.Add($tabControl)
    # Ensure headers remain visible after control tree built
    $grid.ColumnHeadersVisible = $true
    $grid.ColumnHeadersHeight = [Math]::Max(30, $grid.ColumnHeadersHeight)
    $grid.Invalidate(); $grid.Refresh()
    # Reassert patches row height once more after layout using exposed variables
    $null = $grid.BeginInvoke([System.Action]{ try { if ($grid.Rows.Count -gt 0) { $pi=$grid.Rows.Count-1; $lh=$grid.Font.Height+6; $mult=([double]($grid.Tag.PatchesHeightMult)); if ($mult -le 0) { $mult=8 } $minPx=([int]($grid.Tag.PatchesMinPx)); if ($minPx -le 0) { $minPx=80 } $mh=[Math]::Max([int]([Math]::Ceiling($mult*$lh)),$minPx); $grid.Rows[$pi].MinimumHeight=$mh; if ($grid.Rows[$pi].Height -lt $mh) { $grid.Rows[$pi].Height=$mh } } } catch {} })
    # Allow ESC key to close this CSVER window
    $resultsForm.KeyPreview = $true
    $resultsForm.Add_KeyDown({ if ($_.KeyCode -eq 'Escape') { $resultsForm.Close() } })

    $copyButton = New-Object System.Windows.Forms.Button
    $copyButton.Text = "Copy Details"
    $copyButton.Location = New-Object System.Drawing.Point(320, 620)
    $copyButton.Size = New-Object System.Drawing.Size(110, 30)
    $copyButton.Anchor = "Bottom"
    $copyButton.BackColor = [System.Drawing.Color]::LightGreen
    $copyButton.Font = New-Object System.Drawing.Font("Arial", 9, [System.Drawing.FontStyle]::Bold)
    $copyButton.Add_Click({ try { [System.Windows.Forms.Clipboard]::SetText($resultsTextBox.Text); $copyButton.Text = "Copied!"; $copyButton.BackColor = [System.Drawing.Color]::Gold; $timer = New-Object System.Windows.Forms.Timer; $timer.Interval = 1500; $timer.Add_Tick({ $copyButton.Text = "Copy Details"; $copyButton.BackColor = [System.Drawing.Color]::LightGreen; $timer.Stop(); $timer.Dispose() }); $timer.Start() } catch { [System.Windows.Forms.MessageBox]::Show("Failed to copy to clipboard: $_", "Copy Error", "OK", "Error") } })
    $resultsForm.Controls.Add($copyButton)

    $copySummaryBtn = New-Object System.Windows.Forms.Button
    $copySummaryBtn.Text = "Copy Summary"
    $copySummaryBtn.Location = New-Object System.Drawing.Point(440, 620)
    $copySummaryBtn.Size = New-Object System.Drawing.Size(140, 30)
    $copySummaryBtn.Anchor = "Bottom"
    $copySummaryBtn.BackColor = [System.Drawing.Color]::LightGreen
    $copySummaryBtn.Font = New-Object System.Drawing.Font("Arial", 9, [System.Drawing.FontStyle]::Bold)
    $copySummaryBtn.Add_Click({
        try {
            $csvLines = $grid.Tag.CsvLines
            if ($null -eq $csvLines -or $csvLines.Count -eq 0) {
                [System.Windows.Forms.MessageBox]::Show('No summary data to copy.', 'Copy Summary', 'OK', 'Information')
            } else {
                [System.Windows.Forms.Clipboard]::SetText(($csvLines -join "`r`n"))
                $copySummaryBtn.Text = 'Copied!'
                $copySummaryBtn.BackColor = [System.Drawing.Color]::Gold
                $t = New-Object System.Windows.Forms.Timer
                $t.Interval = 1500
                $t.Add_Tick({ $copySummaryBtn.Text = 'Copy Summary'; $copySummaryBtn.BackColor = [System.Drawing.Color]::LightGreen; $t.Stop(); $t.Dispose() })
                $t.Start()
            }
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Failed to copy summary: $_", 'Copy Error', 'OK', 'Error')
        }
    })
    $resultsForm.Controls.Add($copySummaryBtn)

    $closeButton = New-Object System.Windows.Forms.Button
    $closeButton.Text = "Close"
    $closeButton.Location = New-Object System.Drawing.Point(480, 620)
    $closeButton.Size = New-Object System.Drawing.Size(80, 30)
    $closeButton.Anchor = "Bottom"
    $closeButton.Add_Click({ $resultsForm.Close() })
    $resultsForm.Controls.Add($closeButton)

    [void]$resultsForm.ShowDialog()
    Log-Message "=== CSVER REPORT OPERATION COMPLETED ==="
}

 

# IIS Recycle Settings report (migrated from button handler)
function Invoke-IISRecycleSettings {
    if (-not (Validate-Servers)) { return }

    $selectedServers = Get-AllSelectedServers
    Log-Message "=== IIS RECYCLE SETTINGS REPORT OPERATION STARTED ==="
    Log-Message "Selected Servers: $($selectedServers -join ', ')"

    # Collect IIS App Pool settings per server
    $allAppPoolData = @{}
    $appPoolNameSet = New-Object System.Collections.Generic.HashSet[string]

    foreach ($server in $selectedServers | Sort-Object) {
        try {
            $res = Invoke-Command -ComputerName $server -ScriptBlock {
                $serverName = ${env:COMPUTERNAME}
                try {
                    Import-Module WebAdministration -ErrorAction Stop
                } catch {
                    return [pscustomobject]@{ Server = $serverName; Error = "WebAdministration module not available" }
                }
                try {
                    $pools = Get-ChildItem IIS:\AppPools | ForEach-Object {
                        $timeHours = 0
                        if ($_.recycling.periodicRestart.time) { $timeHours = $_.recycling.periodicRestart.time.TotalHours }
                        $times = ''
                        try {
                            $times = ($_.recycling.periodicRestart.schedule.collection | ForEach-Object { $_.value }) -join ','
                        } catch { $times = '' }
                        [PSCustomObject]@{
                            Name            = $_.Name
                            PrivateMemoryKB = $_.recycling.periodicRestart.privateMemory
                            Requests        = $_.recycling.periodicRestart.requests
                            SpecificTime    = $times
                            TimeHours       = $timeHours
                        }
                    }
                } catch {
                    return [pscustomobject]@{ Server = $serverName; Error = "Failed to enumerate AppPools: $($_.Exception.Message)" }
                }
                return [pscustomobject]@{ Server = $serverName; Pools = $pools }
            } -ErrorAction Stop

            $allAppPoolData[$server] = @{}
            if ($null -ne $res -and $null -ne $res.Error -and ($res.Error -ne "")) {
                Log-Message "IIS data error on ${server}: $($res.Error)"
            } elseif ($null -ne $res -and $null -ne $res.Pools) {
                $poolCount = 0
                foreach ($p in $res.Pools) {
                    $name = [string]$p.Name
                    if (-not [string]::IsNullOrWhiteSpace($name)) {
                        [void]$appPoolNameSet.Add($name)
                        $mem = [long]0
                        [long]::TryParse([string]$p.PrivateMemoryKB, [ref]$mem) | Out-Null
                        $req = [int]0
                        [int]::TryParse([string]$p.Requests, [ref]$req) | Out-Null
                        $hrs = [double]0
                        [double]::TryParse([string]$p.TimeHours, [ref]$hrs) | Out-Null
                        $allAppPoolData[$server][$name] = [PSCustomObject]@{
                            AppPoolName     = $name
                            PrivateMemoryKB = $mem
                            Requests        = $req
                            SpecificTime    = [string]$p.SpecificTime
                            TimeHours       = $hrs
                        }
                        $poolCount++
                    }
                }
                Log-Message "Retrieved $poolCount app pool(s) from ${server}"
            } else {
                Log-Message "No Pools or Error returned from ${server}"
            }
        } catch {
            Log-Message "ERROR retrieving IIS settings from $server - $_"
            $allAppPoolData[$server] = @{}
        }
    }

    # Build consolidated, unique app pool name list from collected data (robust across PS versions)
    $allAppPoolNames = @()
    foreach ($server in $selectedServers | Sort-Object) {
        if ($allAppPoolData.ContainsKey($server) -and $allAppPoolData[$server].Count -gt 0) {
            $allAppPoolNames += $allAppPoolData[$server].Keys
        }
    }
    $allAppPoolNames = $allAppPoolNames | Sort-Object -Unique
    $report = @()

    # Consolidated report header
    $report += "=== IIS APP POOL RECYCLE SETTINGS REPORT ==="
    $report += "Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    $report += "Servers: $($selectedServers.Count) selected"
    $report += "Total Unique App Pools Found: $($allAppPoolNames.Count)"
    $report += ""

    # TABLE 1: PRIVATE MEMORY LIMIT
    $report += "=== TABLE 1: PRIVATE MEMORY LIMIT (KB) ==="
    $report += ""
    $headerRow = "AppPoolName".PadRight(30)
    foreach ($server in $selectedServers | Sort-Object) { $headerRow += $server.PadLeft(15) }
    $report += $headerRow
    $report += (New-Object string('-', $headerRow.Length))
    foreach ($appPoolName in $allAppPoolNames) {
        $dataRow = $appPoolName.PadRight(30)
        foreach ($server in $selectedServers | Sort-Object) {
            $value = "N/A"
            if ($allAppPoolData.ContainsKey($server) -and $allAppPoolData[$server].ContainsKey($appPoolName)) {
                $memRaw = $allAppPoolData[$server][$appPoolName].PrivateMemoryKB
                $mem = 0
                [int64]::TryParse([string]$memRaw, [ref]$mem) | Out-Null
                $value = if ($mem -gt 0) { "{0:N0}" -f $mem } else { "0" }
            }
            $dataRow += $value.PadLeft(15)
        }
        $report += $dataRow
    }
    $report += ""
    $report += ""

    # TABLE 2: REQUEST LIMITS
    $report += "=== TABLE 2: REQUEST LIMITS ==="
    $report += ""
    $headerRow = "AppPoolName".PadRight(30)
    foreach ($server in $selectedServers | Sort-Object) { $headerRow += $server.PadLeft(15) }
    $report += $headerRow
    $report += (New-Object string('-', $headerRow.Length))
    foreach ($appPoolName in $allAppPoolNames) {
        $dataRow = $appPoolName.PadRight(30)
        foreach ($server in $selectedServers | Sort-Object) {
            $value = "N/A"
            if ($allAppPoolData.ContainsKey($server) -and $allAppPoolData[$server].ContainsKey($appPoolName)) {
                $reqRaw = $allAppPoolData[$server][$appPoolName].Requests
                $req = 0
                [int]::TryParse([string]$reqRaw, [ref]$req) | Out-Null
                $value = if ($req -gt 0) { "{0:N0}" -f $req } else { "0" }
            }
            $dataRow += $value.PadLeft(15)
        }
        $report += $dataRow
    }
    $report += ""
    $report += ""

    # TABLE 3: SPECIFIC RECYCLE TIMES
    $report += "=== TABLE 3: SPECIFIC RECYCLE TIMES ==="
    $report += ""
    $headerRow = "AppPoolName".PadRight(30)
    foreach ($server in $selectedServers | Sort-Object) { $headerRow += $server.PadLeft(15) }
    $report += $headerRow
    $report += (New-Object string('-', $headerRow.Length))
    foreach ($appPoolName in $allAppPoolNames) {
        $dataRow = $appPoolName.PadRight(30)
        foreach ($server in $selectedServers | Sort-Object) {
            $value = "N/A"
            if ($allAppPoolData.ContainsKey($server) -and $allAppPoolData[$server].ContainsKey($appPoolName)) {
                $timeValue = $allAppPoolData[$server][$appPoolName].SpecificTime
                if ([string]::IsNullOrWhiteSpace($timeValue)) { $value = "None" } else { $value = $timeValue }
            }
            $dataRow += $value.PadLeft(15)
        }
        $report += $dataRow
    }
    $report += ""
    $report += ""

    # TABLE 4: TIME-BASED RECYCLE (HOURS)
    $report += "=== TABLE 4: TIME-BASED RECYCLE (HOURS) ==="
    $report += ""
    $headerRow = "AppPoolName".PadRight(30)
    foreach ($server in $selectedServers | Sort-Object) { $headerRow += $server.PadLeft(15) }
    $report += $headerRow
    $report += (New-Object string('-', $headerRow.Length))
    foreach ($appPoolName in $allAppPoolNames) {
        $dataRow = $appPoolName.PadRight(30)
        foreach ($server in $selectedServers | Sort-Object) {
            $value = "N/A"
            if ($allAppPoolData.ContainsKey($server) -and $allAppPoolData[$server].ContainsKey($appPoolName)) {
                $hrsRaw = $allAppPoolData[$server][$appPoolName].TimeHours
                $hrs = 0.0
                [double]::TryParse([string]$hrsRaw, [ref]$hrs) | Out-Null
                $value = if ($hrs -gt 0) { $hrs.ToString("F1") } else { "0" }
            }
            $dataRow += $value.PadLeft(15)
        }
        $report += $dataRow
    }
    $report += ""
    $report += ""

    # DETAILED RESULTS BY SERVER
    $report += "=== DETAILED RESULTS BY SERVER ==="
    $report += ""
    foreach ($server in $selectedServers | Sort-Object) {
        $report += "===== Results for ${server} ====="
        if ($allAppPoolData.ContainsKey($server) -and $allAppPoolData[$server].Count -gt 0) {
            $report += ""
            $report += "AppPoolName".PadRight(35) + "PrivateMemoryKB".PadLeft(15) + "Requests".PadLeft(12) + "SpecificTime".PadLeft(15) + "TimeHours".PadLeft(12)
            $report += (New-Object string('-', 89))
            $serverAppPools = $allAppPoolData[$server].Keys | Sort-Object
            foreach ($appPoolName in $serverAppPools) {
                $appPool = $allAppPoolData[$server][$appPoolName]
                $memVal = 0; [int64]::TryParse([string]$appPool.PrivateMemoryKB, [ref]$memVal) | Out-Null
                $reqVal = 0; [int]::TryParse([string]$appPool.Requests, [ref]$reqVal) | Out-Null
                $hrsVal = 0.0; [double]::TryParse([string]$appPool.TimeHours, [ref]$hrsVal) | Out-Null
                $memoryDisplay = if ($memVal -gt 0) { "{0:N0}" -f $memVal } else { "0" }
                $requestsDisplay = if ($reqVal -gt 0) { "{0:N0}" -f $reqVal } else { "0" }
                $timeDisplay = if ([string]::IsNullOrWhiteSpace([string]$appPool.SpecificTime)) { "None" } else { [string]$appPool.SpecificTime }
                $hoursDisplay = if ($hrsVal -gt 0) { $hrsVal.ToString("F1") } else { "0" }
                $line = $appPool.AppPoolName.PadRight(35) + $memoryDisplay.PadLeft(15) + $requestsDisplay.PadLeft(12) + $timeDisplay.PadLeft(15) + $hoursDisplay.PadLeft(12)
                $report += $line
            }
        } else {
            $report += "No data retrieved or connection failed (check errors in log)"
        }
        $report += ""
        $report += ""
    }

    # Create results window
    $resultsForm = New-Object System.Windows.Forms.Form
    $resultsForm.Text = "IIS App Pool Recycle Settings Report - $($selectedServers.Count) Servers"
    $resultsForm.Size = New-Object System.Drawing.Size(1500, 700)
    $resultsForm.StartPosition = "CenterParent"
    $resultsForm.FormBorderStyle = "Sizable"

    $resultsTextBox = New-Object System.Windows.Forms.TextBox
    $resultsTextBox.Location = New-Object System.Drawing.Point(10, 10)
    $resultsTextBox.Size = New-Object System.Drawing.Size(1460, 600)
    $resultsTextBox.Multiline = $true
    $resultsTextBox.ScrollBars = "Vertical"
    $resultsTextBox.Font = New-Object System.Drawing.Font("Consolas", 9)
    $resultsTextBox.WordWrap = $false
    $resultsTextBox.ReadOnly = $true
    $resultsTextBox.Text = $report -join "`r`n"
    $resultsTextBox.SelectionStart = 0
    $resultsTextBox.SelectionLength = 0
    $resultsTextBox.Anchor = "Top, Left, Bottom, Right"
    $resultsForm.Controls.Add($resultsTextBox)

    $copyButton = New-Object System.Windows.Forms.Button
    $copyButton.Text = "Copy to Clipboard"
    $copyButton.Location = New-Object System.Drawing.Point(450, 620)
    $copyButton.Size = New-Object System.Drawing.Size(120, 30)
    $copyButton.Anchor = "Bottom"
    $copyButton.BackColor = [System.Drawing.Color]::LightGreen
    $copyButton.Font = New-Object System.Drawing.Font("Arial", 9, [System.Drawing.FontStyle]::Bold)
    $copyButton.Add_Click({
        try {
            [System.Windows.Forms.Clipboard]::SetText($resultsTextBox.Text)
            $copyButton.Text = "Copied!"
            $copyButton.BackColor = [System.Drawing.Color]::Gold
            $timer = New-Object System.Windows.Forms.Timer
            $timer.Interval = 1500
            $timer.Add_Tick({
                $copyButton.Text = "Copy to Clipboard"
                $copyButton.BackColor = [System.Drawing.Color]::LightGreen
                $timer.Stop()
                $timer.Dispose()
            })
            $timer.Start()
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Failed to copy to clipboard: $_", "Copy Error", "OK", "Error")
        }
    })
    $resultsForm.Controls.Add($copyButton)

    $closeButton = New-Object System.Windows.Forms.Button
    $closeButton.Text = "Close"
    $closeButton.Location = New-Object System.Drawing.Point(580, 620)
    $closeButton.Size = New-Object System.Drawing.Size(80, 30)
    $closeButton.Anchor = "Bottom"
    $closeButton.Add_Click({ $resultsForm.Close() })
    $resultsForm.Controls.Add($closeButton)

    [void]$resultsForm.ShowDialog()

    Log-Message "=== IIS RECYCLE SETTINGS REPORT OPERATION COMPLETED ==="
}

# COMMON TASKS UI: label, dropdown, run button
$commonTasksLabel = New-Object System.Windows.Forms.Label
$commonTasksLabel.Text = "COMMON TASKS"
$commonTasksLabel.Location = New-Object System.Drawing.Point(821, 760)
$commonTasksLabel.Size = New-Object System.Drawing.Size(200, 18)
$commonTasksLabel.Font = New-Object System.Drawing.Font("Arial", 8, [System.Drawing.FontStyle]::Bold)
$commonTasksLabel.ForeColor = [System.Drawing.Color]::Purple
$commonTasksLabel.Anchor = "Bottom, Right"
$form.Controls.Add($commonTasksLabel)

$commonTasksDropdown = New-Object System.Windows.Forms.ComboBox
$commonTasksDropdown.Location = New-Object System.Drawing.Point(821, 780)
$commonTasksDropdown.Size = New-Object System.Drawing.Size(220, 21)
$commonTasksDropdown.DropDownStyle = "DropDownList"
$commonTasksDropdown.Anchor = "Bottom, Right"
[void]$commonTasksDropdown.Items.AddRange((@(
    "AdminCommand status",
    "AdminTools Not Working",
    "CSVER REPORT",
    "DICOM Node - Apply Changes (DAs)",
    "DICOM Node - Make Visible VMotion (Portals)",
    "Disable MEMO",
    "FlexLM - Temp License Reset",
    "Enable MEMO",
    "Holy3",
    "IIS Recycle Settings Report",
    "Java Service Info",
    "Kick it - HARD",
    "Orchestrator_Normalization Svc Restart",
    "UPTIME REPORT"
) | Sort-Object))
$commonTasksDropdown.SelectedIndex = 0
$form.Controls.Add($commonTasksDropdown)

$commonTasksRunButton = New-Object System.Windows.Forms.Button
$commonTasksRunButton.Text = "RUN TASK"
$commonTasksRunButton.Location = New-Object System.Drawing.Point(1048, 775)
$commonTasksRunButton.Size = New-Object System.Drawing.Size(110, 30)
$commonTasksRunButton.Anchor = "Bottom, Right"
$commonTasksRunButton.BackColor = [System.Drawing.Color]::LightGray
$commonTasksRunButton.ForeColor = [System.Drawing.Color]::Purple
$commonTasksRunButton.Font = New-Object System.Drawing.Font("Arial", 8, [System.Drawing.FontStyle]::Bold)
$commonTasksRunButton.Add_Click({
    $task = $commonTasksDropdown.SelectedItem
    if (-not $task) { [System.Windows.Forms.MessageBox]::Show("Please select a task.", "No Task Selected", "OK", "Information"); return }

    # Tasks that don't require confirmation
    $skipConfirm = @('AdminCommand status','CSVER REPORT','IIS Recycle Settings Report','UPTIME REPORT','Java Service Info')
    if ($skipConfirm -notcontains $task) {
        $servers = Get-AllSelectedServers
        if (-not $servers -or $servers.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Please select at least one server before running this task.", "No Servers Selected", "OK", "Information"); return
        }
        $preview = Build-CommonTaskPreview -TaskName $task -Servers $servers
        # Confirmation dialog with preview of commands
        $cf = New-Object System.Windows.Forms.Form
        $cf.Text = "Confirm: $task"
        $cf.Size = New-Object System.Drawing.Size(900, 600)
        $cf.StartPosition = 'CenterParent'
        $cf.FormBorderStyle = 'Sizable'

        $lbl = New-Object System.Windows.Forms.Label
        $lbl.Text = "This will run the following commands. Are you sure you want to proceed?"
        $lbl.AutoSize = $true
        $lbl.Location = New-Object System.Drawing.Point(10, 10)

        $txt = New-Object System.Windows.Forms.TextBox
        $txt.Multiline = $true
        $txt.ScrollBars = 'Vertical'
        $txt.Font = New-Object System.Drawing.Font('Consolas', 9)
        $txt.WordWrap = $false
        $txt.ReadOnly = $true
        $txt.Location = New-Object System.Drawing.Point(10, 35)
        $txt.Size = New-Object System.Drawing.Size(860, 480)
        $txt.Anchor = 'Top, Left, Bottom, Right'
    $txt.Text = $preview
    $txt.SelectionStart = 0
    $txt.SelectionLength = 0

    # Allow ESC to close confirmation dialog
    $cf.KeyPreview = $true
    $cf.Add_KeyDown({ if ($_.KeyCode -eq 'Escape') { $cf.DialogResult = [System.Windows.Forms.DialogResult]::Cancel; $cf.Close() } })

        $btnProceed = New-Object System.Windows.Forms.Button
        $btnProceed.Text = 'Proceed'
        $btnProceed.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $btnProceed.Location = New-Object System.Drawing.Point(680, 525)
        $btnProceed.Size = New-Object System.Drawing.Size(85, 28)
        $btnProceed.Anchor = 'Bottom, Right'

        $btnCancel = New-Object System.Windows.Forms.Button
        $btnCancel.Text = 'Cancel'
        $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        $btnCancel.Location = New-Object System.Drawing.Point(775, 525)
        $btnCancel.Size = New-Object System.Drawing.Size(85, 28)
        $btnCancel.Anchor = 'Bottom, Right'

        $cf.AcceptButton = $btnProceed
        $cf.CancelButton = $btnCancel
        [void]$cf.Controls.Add($lbl)
        [void]$cf.Controls.Add($txt)
        [void]$cf.Controls.Add($btnProceed)
        [void]$cf.Controls.Add($btnCancel)

        if ($cf.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }
    }
    switch ($task) {
    "AdminCommand status" { Invoke-AdminStatus }
        "AdminTools Not Working" { Invoke-AdminToolsNotWorking }
        "Kick it - HARD" { Invoke-KickItHard }
    "Orchestrator_Normalization Svc Restart" { Invoke-OrchestratorNormalizationSvcRestart }
        "IIS Recycle Settings Report" { Invoke-IISRecycleSettings }
        "CSVER REPORT" { Invoke-CsverReport }
        "DICOM Node - Apply Changes (DAs)" { Invoke-DicomApplyChanges }
        "DICOM Node - Make Visible VMotion (Portals)" { Invoke-PortalMakeVisibleVMotion }
        "Disable MEMO" { Invoke-DisableMemo }
    "FlexLM - Temp License Reset" { Invoke-FlexLMTempLicenseReset }
        "Enable MEMO" { Invoke-EnableMemo }
        "Holy3"        { Invoke-IISHoly3 }
        "Java Service Info" { Invoke-JavaInfo }
        "UPTIME REPORT"{ Invoke-UptimeReport }
        default { [System.Windows.Forms.MessageBox]::Show("Please select a task.", "No Task Selected", "OK", "Information") }
    }
})
$form.Controls.Add($commonTasksRunButton)

# Remove REVIEW TASK button; align RUN TASK next to the dropdown
# Hard-code RUN TASK position and size (top and height match Edit Svr/Svc)
$commonTasksRunButton.Location = New-Object System.Drawing.Point(1048, 775)
$commonTasksRunButton.Size = New-Object System.Drawing.Size(90, 30)

# QUERY CFG UI: label, dropdown, and path button (to the left of COMMON TASKS)
$queryCfgLabel = New-Object System.Windows.Forms.Label
$queryCfgLabel.Text = 'QUERY CFG'
$queryCfgLabel.Location = New-Object System.Drawing.Point(600, 760)
$queryCfgLabel.Size = New-Object System.Drawing.Size(200, 18)
$queryCfgLabel.Font = New-Object System.Drawing.Font('Arial', 8, [System.Drawing.FontStyle]::Bold)
$queryCfgLabel.ForeColor = [System.Drawing.Color]::Navy
$queryCfgLabel.Anchor = 'Bottom, Right'
$form.Controls.Add($queryCfgLabel)

$queryCfgDropdown = New-Object System.Windows.Forms.ComboBox
$queryCfgDropdown.Location = New-Object System.Drawing.Point(600, 780)
$queryCfgDropdown.Size = New-Object System.Drawing.Size(105, 21)
$queryCfgDropdown.DropDownStyle = 'DropDownList'
$queryCfgDropdown.Anchor = 'Bottom, Right'
[void]$queryCfgDropdown.Items.AddRange(@('get_value','get_value_list','get_sub_list','is_subtree_exist'))
$queryCfgDropdown.SelectedIndex = 0
$form.Controls.Add($queryCfgDropdown)

$queryCfgButton = New-Object System.Windows.Forms.Button
$queryCfgButton.Text = 'Specify Path'
$queryCfgButton.Location = New-Object System.Drawing.Point(710, 775)
$queryCfgButton.Size = New-Object System.Drawing.Size(100, 30)
$queryCfgButton.Anchor = 'Bottom, Right'
$queryCfgButton.BackColor = [System.Drawing.Color]::LightGray
$queryCfgButton.ForeColor = [System.Drawing.Color]::Navy
$queryCfgButton.Font = New-Object System.Drawing.Font('Arial', 8, [System.Drawing.FontStyle]::Bold)
$queryCfgButton.Add_Click({
    # Prompt for full CFG path
    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = 'Enter CFG Path'
    $dlg.Size = New-Object System.Drawing.Size(700, 160)
    $dlg.StartPosition = 'CenterParent'
    $dlg.FormBorderStyle = 'FixedDialog'
    $dlg.MaximizeBox = $false
    $dlg.MinimizeBox = $false
    $dlg.KeyPreview = $true
    $dlg.Add_KeyDown({ if ($_.KeyCode -eq 'Escape') { $dlg.DialogResult = [System.Windows.Forms.DialogResult]::Cancel; $dlg.Close() } })
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = 'ENTER THE FULL CFG PATH (e.g., imaginet\\system\\...)'
    $lbl.AutoSize = $true
    $lbl.Location = New-Object System.Drawing.Point(10, 10)
    $tb = New-Object System.Windows.Forms.TextBox
    $tb.Location = New-Object System.Drawing.Point(10, 35)
    $tb.Size = New-Object System.Drawing.Size(660, 22)
    $tb.Anchor = 'Top, Left, Right'
    $ok = New-Object System.Windows.Forms.Button
    $ok.Text = 'OK'
    $ok.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $ok.Location = New-Object System.Drawing.Point(495, 70)
    $ok.Size = New-Object System.Drawing.Size(80, 25)
    $ok.Anchor = 'Bottom, Right'
    $cancel = New-Object System.Windows.Forms.Button
    $cancel.Text = 'Cancel'
    $cancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $cancel.Location = New-Object System.Drawing.Point(590, 70)
    $cancel.Size = New-Object System.Drawing.Size(80, 25)
    $cancel.Anchor = 'Bottom, Right'
    $dlg.AcceptButton = $ok
    $dlg.CancelButton = $cancel
    [void]$dlg.Controls.Add($lbl)
    [void]$dlg.Controls.Add($tb)
    [void]$dlg.Controls.Add($ok)
    [void]$dlg.Controls.Add($cancel)
    if ($dlg.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }
    $path = $tb.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($path)) { [System.Windows.Forms.MessageBox]::Show('Path cannot be empty.','Validation','OK','Warning'); return }

    # Determine command from dropdown
    $cmd = $queryCfgDropdown.SelectedItem
    if (-not $cmd) { [System.Windows.Forms.MessageBox]::Show('Please select a Query CFG command.','No Command','OK','Information'); return }

    # Gather selected servers
    $servers = Get-AllSelectedServers
    if (-not $servers -or $servers.Count -eq 0) { [System.Windows.Forms.MessageBox]::Show('Please select at least one server.','No Servers','OK','Information'); return }

    # Log start and intent
    try {
        Log-Message "=== QUERY CFG OPERATION STARTED ==="
        Log-Message "Query CFG Command: $cmd | Path: $path"
        Log-Message ("Selected Servers: {0}" -f ($servers -join ', '))
        Log-Message "Results window will open after all servers have returned results."
    } catch {}

    # Execute tool_cfg on each server
    $resultsByServer = @{}
    foreach ($server in $servers) {
        try {
            $output = Invoke-Command -ComputerName $server -ScriptBlock {
                param($c,$p)
                $cmdLine = "tool_cfg -c $c -p `"$p`""
                try {
                    $res = Invoke-Expression $cmdLine 2>&1 | Out-String
                    return $res
                } catch {
                    return "ERROR executing: $cmdLine`n$_"
                }
            } -ArgumentList $cmd, $path -ErrorAction Stop
            $resultsByServer[$server] = [string]$output
        } catch {
            $resultsByServer[$server] = "ERROR contacting $server - $_"
        }
    }

    # Build a results table with server names in the first row and split multi-line results into separate rows
    $rf = New-Object System.Windows.Forms.Form
    $rf.Text = "Query CFG Results"
    $rf.Size = New-Object System.Drawing.Size(1200, 600)
    $rf.StartPosition = 'CenterParent'
    $rf.FormBorderStyle = 'Sizable'
    $rf.KeyPreview = $true
    $rf.Add_KeyDown({ if ($_.KeyCode -eq 'Escape') { $rf.Close() } })

    $lblHdr = New-Object System.Windows.Forms.Label
    $lblHdr.Text = "Path: $path"
    $lblHdr.AutoSize = $true
    $lblHdr.Font = New-Object System.Drawing.Font('Arial', 9, [System.Drawing.FontStyle]::Bold)
    $lblHdr.Location = New-Object System.Drawing.Point(10, 10)

    $grid = New-Object System.Windows.Forms.DataGridView
    $grid.Location = New-Object System.Drawing.Point(10, 35)
    $grid.Size = New-Object System.Drawing.Size(1160, 510)
    $grid.Anchor = 'Top, Left, Bottom, Right'
    $grid.ReadOnly = $true
    $grid.AllowUserToAddRows = $false
    $grid.AllowUserToDeleteRows = $false
    $grid.RowHeadersVisible = $false
    $grid.ColumnHeadersVisible = $false
    # Use manual sizing to allow interactive resize; avoids expensive autosize passes
    $grid.AutoSizeColumnsMode = 'None'
    $grid.AutoSizeRowsMode = 'None'
    $grid.AllowUserToResizeColumns = $true
    # Tag used for width sync reentrancy guard
    $grid.Tag = @{ WidthSync = $false }
    $grid.DefaultCellStyle.WrapMode = 'False'

    # Preprocess results into per-server line arrays (trim trailing empty lines) using a growable list
    $perServerLines = @{}
    $maxLines = 0
    foreach ($s in $servers) {
        $raw = if ($resultsByServer.ContainsKey($s)) { [string]$resultsByServer[$s] } else { '' }
        $list = New-Object System.Collections.Generic.List[string]
        foreach ($ln in ($raw -split "`r?`n")) { $list.Add([string]$ln) }
        while ($list.Count -gt 0 -and [string]::IsNullOrWhiteSpace($list[$list.Count - 1])) { $list.RemoveAt($list.Count - 1) }
        $perServerLines[$s] = $list
        if ($list.Count -gt $maxLines) { $maxLines = $list.Count }
    }

    # Build DataTable with one column per server; headers hidden, so the first row will hold server names
    $dt = New-Object System.Data.DataTable
    foreach ($s in $servers) { [void]$dt.Columns.Add([string]$s) }

    # Row 0: server names
    $rHeader = $dt.NewRow()
    foreach ($s in $servers) { $rHeader[$s] = $s }
    [void]$dt.Rows.Add($rHeader)

    # Subsequent rows: line i from each server
    for ($i = 0; $i -lt $maxLines; $i++) {
        $r = $dt.NewRow()
        foreach ($s in $servers) {
            $cells = $perServerLines[$s]
            $r[$s] = if ($i -lt $cells.Count) { [string]$cells[$i] } else { '' }
        }
        [void]$dt.Rows.Add($r)
    }

    $grid.DataSource = $dt
    # Reassert header hiding after binding to be safe and style the top row (server names)
    $grid.Add_DataBindingComplete({
        $grid.ColumnHeadersVisible = $false
        if ($grid.Rows.Count -gt 0) {
            $grid.Rows[0].DefaultCellStyle.BackColor = [System.Drawing.Color]::Gainsboro
            $grid.Rows[0].DefaultCellStyle.Font = New-Object System.Drawing.Font($grid.Font, [System.Drawing.FontStyle]::Bold)
            $grid.Rows[0].DefaultCellStyle.Alignment = [System.Windows.Forms.DataGridViewContentAlignment]::MiddleCenter
        }
    })
    if ($grid.Rows.Count -gt 0) {
        $grid.Rows[0].DefaultCellStyle.BackColor = [System.Drawing.Color]::Gainsboro
        $grid.Rows[0].DefaultCellStyle.Font = New-Object System.Drawing.Font($grid.Font, [System.Drawing.FontStyle]::Bold)
        $grid.Rows[0].DefaultCellStyle.Alignment = [System.Windows.Forms.DataGridViewContentAlignment]::MiddleCenter
    }
    $grid.ColumnHeadersVisible = $false; $grid.Invalidate(); $grid.Refresh()

    # When any column is resized, apply the same width to all columns
    $grid.Add_ColumnWidthChanged({
        param($s, $e)
        if ($null -eq $e -or $null -eq $e.Column) { return }
        if ($grid.Tag -and $grid.Tag.WidthSync) { return }
        try {
            if (-not $grid.Tag) { $grid.Tag = @{ WidthSync = $false } }
            $grid.Tag.WidthSync = $true
            $w = [int]$e.Column.Width
            foreach ($c in $grid.Columns) { if ($null -ne $c -and $c.Width -ne $w) { $c.Width = $w } }
        } finally {
            if ($grid.Tag) { $grid.Tag.WidthSync = $false }
        }
    })

    # Context menu for copying a single cell value
    $cms = New-Object System.Windows.Forms.ContextMenuStrip
    $miCopyCell = New-Object System.Windows.Forms.ToolStripMenuItem
    $miCopyCell.Text = 'Copy Cell'
    $miCopyCell.Add_Click({ if ($grid.CurrentCell -ne $null) { [System.Windows.Forms.Clipboard]::SetText([string]$grid.CurrentCell.Value) } })
    [void]$cms.Items.Add($miCopyCell)
    $grid.ContextMenuStrip = $cms

    # Copy All button (tab-delimited: first row = server names; following rows = per-line results)
    $btnCopyAll = New-Object System.Windows.Forms.Button
    $btnCopyAll.Text = 'Copy All'
    $btnCopyAll.Location = New-Object System.Drawing.Point(100, 555)
    $btnCopyAll.Size = New-Object System.Drawing.Size(90, 25)
    $btnCopyAll.Anchor = 'Bottom, Left'
    $btnCopyAll.BackColor = [System.Drawing.Color]::LightGreen
    $btnCopyAll.Font = New-Object System.Drawing.Font('Arial', 8, [System.Drawing.FontStyle]::Bold)
    $btnCopyAll.Add_Click({
        try {
            $sb = New-Object System.Text.StringBuilder
            # header
            [void]$sb.AppendLine(($servers -join "`t"))
            # rows
            for ($i = 0; $i -lt $maxLines; $i++) {
                $rowVals = @()
                foreach ($s in $servers) {
                    $arr = $perServerLines[$s]
                    $val = if ($i -lt $arr.Count) { [string]$arr[$i] } else { '' }
                    $rowVals += $val
                }
                [void]$sb.AppendLine(($rowVals -join "`t"))
            }
            [System.Windows.Forms.Clipboard]::SetText($sb.ToString())
            $btnCopyAll.Text = 'Copied!'; $btnCopyAll.BackColor = [System.Drawing.Color]::Gold
            $t = New-Object System.Windows.Forms.Timer; $t.Interval = 1500
            $t.Add_Tick({ $btnCopyAll.Text = 'Copy All'; $btnCopyAll.BackColor = [System.Drawing.Color]::LightGreen; $t.Stop(); $t.Dispose() })
            $t.Start()
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Failed to copy: $_", 'Copy Error', 'OK', 'Error')
        }
    })

    # Close button for convenience (ESC also closes)
    $btnClose = New-Object System.Windows.Forms.Button
    $btnClose.Text = 'Close'
    $btnClose.Location = New-Object System.Drawing.Point(10, 555)
    $btnClose.Size = New-Object System.Drawing.Size(80, 25)
    $btnClose.Anchor = 'Bottom, Left'
    $btnClose.Add_Click({ $rf.Close() })

    $rf.Controls.Add($lblHdr)
    $rf.Controls.Add($grid)
    $rf.Controls.Add($btnClose)
    $rf.Controls.Add($btnCopyAll)

    try { Log-Message "=== QUERY CFG OPERATION COMPLETED: displaying results ===" } catch {}
    [void]$rf.ShowDialog()
})
$form.Controls.Add($queryCfgButton)

## EXIT BUTTON 
$exitButton = New-Object System.Windows.Forms.Button
$exitButton.Text = "EXIT"
$exitButton.Location = New-Object System.Drawing.Point(1240, 775)
$exitButton.Size = New-Object System.Drawing.Size(80, 30)
$exitButton.Anchor = "Bottom, Right"
$exitButton.Add_Click({
    $form.Close()
})
$form.Controls.Add($exitButton)

# Initialize with a welcome message
Log-Message "Vue PACS Server Management GUI v$AppVersion - Ready for operations"
Log-Message "Select servers and components, then use individual section buttons or RESTART ALL"
Log-Message "Use EDIT button to modify servers, services, etc."
Log-Message "Execute system5 commands with drop down menus"
Log-Message "Commands (except reboot) will open in separate PowerShell windows"

# Show the form
$form.ShowDialog()