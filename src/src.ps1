# MIT License
#
# Copyright (c) 2025 iappyx
#
# Permission is hereby granted, free of charge, to any person obtaining a copy
# of this software and associated documentation files (the "Software"), to deal
# in the Software without restriction, including without limitation the rights
# to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
# copies of the Software, and to permit persons to whom the Software is
# furnished to do so, subject to the following conditions:
# 
# The above copyright notice and this permission notice shall be included in all
# copies or substantial portions of the Software.
# 
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
# OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
# SOFTWARE.

Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
using System.Text;

public class KeyboardListener {
    [DllImport("user32.dll")]
    public static extern short GetAsyncKeyState(int vKey);
}

public class WindowHelper {
    [DllImport("user32.dll")]
    public static extern IntPtr GetForegroundWindow();

    [DllImport("user32.dll")]
    public static extern int GetWindowThreadProcessId(IntPtr hWnd, out int lpdwProcessId);

    [DllImport("user32.dll")]
    public static extern int ShowWindow(IntPtr hWnd, int nCmdShow);

    [DllImport("user32.dll")]
    public static extern bool SetForegroundWindow(IntPtr hWnd);
}
"@ -Language CSharp

$global:hWnd = [WindowHelper]::GetForegroundWindow()
$global:keyMap = @{} 

$instrumentaKeysVersion = "0.18"

Write-Host "██╗███╗   ██╗███████╗████████╗██████╗ ██╗   ██╗███╗   ███╗███████╗███╗   ██╗████████╗ █████╗ "
Write-Host "██║████╗  ██║██╔════╝╚══██╔══╝██╔══██╗██║   ██║████╗ ████║██╔════╝████╗  ██║╚══██╔══╝██╔══██╗"
Write-Host "██║██╔██╗ ██║███████╗   ██║   ██████╔╝██║   ██║██╔████╔██║█████╗  ██╔██╗ ██║   ██║   ███████║"
Write-Host "██║██║╚██╗██║╚════██║   ██║   ██╔══██╗██║   ██║██║╚██╔╝██║██╔══╝  ██║╚██╗██║   ██║   ██╔══██║"
Write-Host "██║██║ ╚████║███████║   ██║   ██║  ██║╚██████╔╝██║ ╚═╝ ██║███████╗██║ ╚████║   ██║   ██║  ██║"
Write-Host "╚═╝╚═╝  ╚═══╝╚══════╝   ╚═╝   ╚═╝  ╚═╝ ╚═════╝ ╚═╝     ╚═╝╚══════╝╚═╝  ╚═══╝   ╚═╝   ╚═╝  ╚═╝"
Write-Host "██╗  ██╗███████╗██╗   ██╗███████╗                                                            "
Write-Host "██║ ██╔╝██╔════╝╚██╗ ██╔╝██╔════╝                                                            "
Write-Host "█████╔╝ █████╗   ╚████╔╝ ███████╗      Keyboard Shortcut Companion (v. $instrumentaKeysVersion)"
Write-Host "██╔═██╗ ██╔══╝    ╚██╔╝  ╚════██║                                                            "
Write-Host "██║  ██╗███████╗   ██║   ███████║                                                            "
Write-Host "╚═╝  ╚═╝╚══════╝   ╚═╝   ╚══════╝                                                            "
Write-Host ""

$shortcuts = @{}
$friendlyShortcuts = @{}
$global:shortcutList = New-Object System.Collections.ArrayList
$csvPath = "shortcuts.csv"

if (-Not (Test-Path $csvPath)) {
    Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - Shortcuts file missing. Creating default configuration..."
    
    $defaultShortcuts = @(
        "Ctrl+Shift+Alt+Q,InstrumentaKeysEditor"
        "Ctrl+Shift+Alt+S,ObjectsSwapPosition"
        "Ctrl+Shift+Alt+L,ObjectsAlignLefts"
        "Ctrl+Shift+Alt+T,ObjectsAlignTops"
        "Ctrl+Shift+Alt+R,ObjectsAlignRights"
        "Ctrl+Shift+Alt+B,ObjectsAlignBottoms"
        "Ctrl+Shift+Alt+E,ObjectsAlignCenters"
        "Ctrl+Shift+Alt+M,ObjectsAlignMiddles"
        "Ctrl+Shift+Alt+H,ObjectsDistributeHorizontally"
        "Ctrl+Shift+Alt+V,ObjectsDistributeVertically"
        "Ctrl+Alt+Left,MoveTableColumnLeft"
        "Ctrl+Alt+Right,MoveTableColumnRight"
        "Ctrl+Alt+Up,MoveTableRowUp"
        "Ctrl+Alt+Down,MoveTableRowDown"
        "Alt+Q,GenerateStickyNote"
    )

    $defaultShortcuts | Out-File -FilePath $csvPath -Encoding UTF8
    Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - Default shortcuts saved in '$csvPath'."
}

function Reload-ShortcutSettings {
    $newShortcutList = New-Object System.Collections.ArrayList

    $shortcuts.Clear()
    $friendlyShortcuts.Clear()

    $global:keyMap = @{}

    if (Test-Path $csvPath) {
        $csvData = Import-Csv -Path $csvPath -Header "Key", "Macro"

        foreach ($entry in $csvData) {
            if ($entry.Key -and $entry.Macro) {
                $keyCombo = $entry.Key -split '\+'
                $virtualKeys = @()

                foreach ($key in $keyCombo) {
                    if (-not $global:keyMap.ContainsKey($key)) {
                        try {
                            switch ($key) {
                                "Ctrl"           { $global:keyMap[$key] = 0x11 }
                                "Shift"          { $global:keyMap[$key] = 0x10 }
                                "Alt"            { $global:keyMap[$key] = 0x12 }
                                "Del"            { $global:keyMap[$key] = 0x2E }
                                "Up"             { $global:keyMap[$key] = 0x26 }
                                "Down"           { $global:keyMap[$key] = 0x28 }
                                "Left"           { $global:keyMap[$key] = 0x25 }
                                "Right"          { $global:keyMap[$key] = 0x27 }
                                "Esc"            { $global:keyMap[$key] = 0x1B }
                                "Enter"          { $global:keyMap[$key] = 0x0D }
                                "Tab"            { $global:keyMap[$key] = 0x09 }
                                "Space"          { $global:keyMap[$key] = 0x20 }
                                "Backspace"      { $global:keyMap[$key] = 0x08 }
                                "PageUp"         { $global:keyMap[$key] = 0x21 }
                                "PageDown"       { $global:keyMap[$key] = 0x22 }
                                "Home"           { $global:keyMap[$key] = 0x24 }
                                "End"            { $global:keyMap[$key] = 0x23 }
                                "Insert"         { $global:keyMap[$key] = 0x2D }
                                "Plus"           { $global:keyMap[$key] = 0x2B }
                                "F1"             { $global:keyMap[$key] = 0x70 }
                                "F2"             { $global:keyMap[$key] = 0x71 }
                                "F3"             { $global:keyMap[$key] = 0x72 }
                                "F4"             { $global:keyMap[$key] = 0x73 }
                                "F5"             { $global:keyMap[$key] = 0x74 }
                                "F6"             { $global:keyMap[$key] = 0x75 }
                                "F7"             { $global:keyMap[$key] = 0x76 }
                                "F8"             { $global:keyMap[$key] = 0x77 }
                                "F9"             { $global:keyMap[$key] = 0x78 }
                                "F10"            { $global:keyMap[$key] = 0x79 }
                                "F11"            { $global:keyMap[$key] = 0x7A }
                                "F12"            { $global:keyMap[$key] = 0x7B }
                                "Num0"           { $global:keyMap[$key] = 0x60 }
                                "Num1"           { $global:keyMap[$key] = 0x61 }
                                "Num2"           { $global:keyMap[$key] = 0x62 }
                                "Num3"           { $global:keyMap[$key] = 0x63 }
                                "Num4"           { $global:keyMap[$key] = 0x64 }
                                "Num5"           { $global:keyMap[$key] = 0x65 }
                                "Num6"           { $global:keyMap[$key] = 0x66 }
                                "Num7"           { $global:keyMap[$key] = 0x67 }
                                "Num8"           { $global:keyMap[$key] = 0x68 }
                                "Num9"           { $global:keyMap[$key] = 0x69 }
                                "NumLock"        { $global:keyMap[$key] = 0x90 }
                                "NumpadDivide"   { $global:keyMap[$key] = 0x6F }
                                "NumpadMultiply" { $global:keyMap[$key] = 0x6A }
                                "NumpadSubtract" { $global:keyMap[$key] = 0x6D }
                                "NumpadAdd"      { $global:keyMap[$key] = 0x6B }
                                "NumpadEnter"    { $global:keyMap[$key] = 0x0D }
                                "CapsLock"       { $global:keyMap[$key] = 0x14 }
                                "ScrollLock"     { $global:keyMap[$key] = 0x91 }
                                "PrintScreen"    { $global:keyMap[$key] = 0x2C }
                                "PauseBreak"     { $global:keyMap[$key] = 0x13 }
                                Default      {
                                    if ([int][char]$key -ge 32 -and [int][char]$key -le 96) {
                                        $global:keyMap[$key] = [int][char]$key
                                    } else {
                                        Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - ERROR: Ignoring unsupported key '$key'."
                                    }
                                }
                            }
                        } catch {
                            Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - ERROR: Failed to process key '$key' in shortcut '$entry.Key'."
                        }
                    }

                    if ($global:keyMap.ContainsKey($key)) {
                        $virtualKeys += $global:keyMap[$key]
                    }
                }

                $macroName = $entry.Macro.Trim()
                if ($macroName -ne "") {
                    $keyIdentifier = $virtualKeys -join " "
                    $shortcuts[$keyIdentifier] = $macroName
                    $friendlyShortcuts[$keyIdentifier] = $entry.Key

                    $newShortcutList.Add([PSCustomObject]@{
                        Shortcut = $entry.Key
                        Macro    = $macroName
                    }) | Out-Null
                }
            } else {
                Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - ERROR: Missing key or macro in CSV entry."
            }
        }
    } else {
        Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - ERROR: CSV file not found!"
        exit
    }

    $global:shortcutList = $newShortcutList

    Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - Shortcut settings loaded"
    Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - Available shortcuts: $($newShortcutList.count)"
    $newShortcutList | Format-Table -AutoSize | Out-Host
}

Reload-ShortcutSettings

function Export-Shortcuts {
    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveFileDialog.InitialDirectory = [System.Environment]::GetFolderPath("Desktop")
    $saveFileDialog.Filter = "CSV Files (*.csv)|*.csv"
    $saveFileDialog.Title = "Save Shortcuts As"

    if ($saveFileDialog.ShowDialog() -eq "OK") {
        $exportPath = $saveFileDialog.FileName
        Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - Exporting shortcuts to: $exportPath"

        $newData = New-Object System.Collections.ArrayList  
        foreach ($row in $grid.Rows) {
            if ($row.Cells[0].Value -and $row.Cells[1].Value) {
                $newData.Add("$($row.Cells[0].Value),$($row.Cells[1].Value)") | Out-Null 
            }
        }

        $newData | Out-File $exportPath -Encoding utf8
        [System.Windows.Forms.MessageBox]::Show("Shortcuts saved successfully!", "Export Complete")
    }
}

function Import-Shortcuts {
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.InitialDirectory = [System.Environment]::GetFolderPath("Desktop")
    $openFileDialog.Filter = "CSV Files (*.csv)|*.csv"
    $openFileDialog.Title = "Select Shortcut File"

    if ($openFileDialog.ShowDialog() -eq "OK") {
        $importPath = $openFileDialog.FileName
        Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - Importing shortcuts from: $importPath"
        
        if (Test-Path $importPath) {
            $csvData = Get-Content -Path $importPath
            $grid.Rows.Clear()  

            foreach ($entry in $csvData) {
                $splitEntry = $entry -split ","
                if ($splitEntry.Count -eq 2) {
                    $grid.Rows.Add($splitEntry[0].Trim(), $splitEntry[1].Trim()) | Out-Null
                }
            }
        } else {
            [System.Windows.Forms.MessageBox]::Show("Error: Could not find the selected file!", "Import Failed")
        }
    }
}

function Import-ShortcutsFromGitHub {
    $repoUrl = "https://api.github.com/repos/iappyx/Instrumenta-Keys/contents/shared-shortcuts/"

    try {
        Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - Fetching CSV files from GitHub..."
        $files = Invoke-RestMethod -Uri $repoUrl
        $csvFiles = $files | Where-Object { $_.name -match "\.csv$" }

        if ($csvFiles.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("No CSV files found in GitHub folder!", "Import Failed")
            return
        }

        $form = New-Object System.Windows.Forms.Form
        $form.Text = "Select CSV File"
        $form.Width = 400
        $form.Height = 300
        $form.StartPosition = "CenterScreen"

        $listBox = New-Object System.Windows.Forms.ListBox
        $listBox.Dock = "Fill"
        foreach ($file in $csvFiles) {
            $listBox.Items.Add($file.name)
        }

        $selectButton = New-Object System.Windows.Forms.Button
        $selectButton.Text = "Import"
        $selectButton.Dock = "Bottom"
        $selectButton.Add_Click({
            $selectedFileName = $listBox.SelectedItem
            if (-not $selectedFileName) {
                [System.Windows.Forms.MessageBox]::Show("Please select a file", "Error")
                return
            }

            $downloadUrl = ($csvFiles | Where-Object { $_.name -eq $selectedFileName }).download_url
            Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - Downloading: $selectedFileName from GitHub..."
            
            $tempCsvPath = "$env:temp\shortcuts_import.csv"
            Invoke-WebRequest -Uri $downloadUrl -OutFile $tempCsvPath

            $csvData = Get-Content -Path $tempCsvPath
            Remove-Item -Path $tempCsvPath -Force 

            $grid.Rows.Clear()
            foreach ($entry in $csvData) {
                $splitEntry = $entry -split ","
                if ($splitEntry.Count -eq 2) {
                    $grid.Rows.Add($splitEntry[0].Trim(), $splitEntry[1].Trim()) | Out-Null
                }
            }

            [System.Windows.Forms.MessageBox]::Show("Imported $selectedFileName successfully!", "Import complete")
            Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - Imported: $selectedFileName"
            $form.Close()
        })

        $form.Controls.Add($listBox)
        $form.Controls.Add($selectButton)

        $form.TopMost = $true
        $exePath = [System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName
        $form.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($exePath)
        $form.ShowDialog()
        
    } catch {
        Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - ERROR: Failed to import shortcuts from GitHub!"
    }
}

function Start-ShortcutEditor {
    Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - Launching shortcut editor..."
    
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Shortcut editor"
    $form.Width = 700
    $form.Height = 400
    $form.StartPosition = "CenterScreen"
    $form.TopMost = $true 

    $mainPanel = New-Object System.Windows.Forms.Panel
    $mainPanel.Dock = "Fill"

    $gridPanel = New-Object System.Windows.Forms.Panel
    $gridPanel.Dock = "Fill"
    $gridPanel.Padding = New-Object System.Windows.Forms.Padding(20)

    $grid = New-Object System.Windows.Forms.DataGridView
    $grid.Dock = "Fill"
    $grid.AutoGenerateColumns = $false

    $colShortcut = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colShortcut.HeaderText = "Shortcut"
    $colShortcut.AutoSizeMode = "Fill"

    $colMacro = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colMacro.HeaderText = "Macro"
    $colMacro.AutoSizeMode = "Fill"

    $grid.Columns.Add($colShortcut) | Out-Null
    $grid.Columns.Add($colMacro) | Out-Null

    if (Test-Path $csvPath) {
        $csvData = Get-Content -Path $csvPath
        foreach ($entry in $($csvData)) {
            $splitEntry = $entry -split ","
            if ($splitEntry.Count -eq 2) {
                $grid.Rows.Add($splitEntry[0].Trim(), $splitEntry[1].Trim()) | Out-Null
            }
        }
    }

    $gridPanel.Controls.Add($grid)
    $buttonPanel = New-Object System.Windows.Forms.Panel
    $buttonPanel.Dock = "Bottom"
    $buttonPanel.Height = 60
   
    $saveButton = New-Object System.Windows.Forms.Button
    $saveButton.Text = "Save Changes"
    $saveButton.Width = 120
    $saveButton.Location = New-Object System.Drawing.Point(410, 5)
    $saveButton.Anchor = "Bottom, Left"

    $saveButton.Add_Click({
        Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - Saving shortcut changes..."
        $newData = New-Object System.Collections.ArrayList  
        
        foreach ($row in $grid.Rows) {
            if ($row.Cells[0].Value -and $row.Cells[1].Value) {
                $newData.Add("$($row.Cells[0].Value),$($row.Cells[1].Value)") | Out-Null 
            }
        }

        Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - Writing shortcuts to CSV file."
        $newData | Out-File $csvPath -Encoding utf8
        
        Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - Closing editor window."
        $form.Close() | Out-Null
        $grid.Dispose() | Out-Null
        Start-Sleep -Milliseconds 500  
        Reload-ShortcutSettings
    })

    $importButton = New-Object System.Windows.Forms.Button
    $importButton.Text = "Import Shortcuts"
    $importButton.Width = 120
    $importButton.Location = New-Object System.Drawing.Point(20, 5)
    $importButton.Anchor = "Bottom, Left"
    $importButton.Add_Click({ Import-Shortcuts })

    $exportButton = New-Object System.Windows.Forms.Button
    $exportButton.Text = "Export Shortcuts"
    $exportButton.Width = 120
    $exportButton.Location = New-Object System.Drawing.Point(150, 5)
    $exportButton.Anchor = "Bottom, Left"
    $exportButton.Add_Click({ Export-Shortcuts })

    $importGitHubButton = New-Object System.Windows.Forms.Button
    $importGitHubButton.Text = "Import from GitHub"
    $importGitHubButton.Width = 120
    $importGitHubButton.Location = New-Object System.Drawing.Point(280, 5)
    $importGitHubButton.Anchor = "Bottom, Left"
    $importGitHubButton.Add_Click({ Import-ShortcutsFromGitHub })

    $buttonPanel.Controls.Add($importGitHubButton)
    $buttonPanel.Controls.Add($importButton)
    $buttonPanel.Controls.Add($exportButton)
    $buttonPanel.Controls.Add($saveButton)
    $mainPanel.Controls.Add($gridPanel)
    $mainPanel.Controls.Add($buttonPanel)

    $form.Controls.Add($mainPanel)
    $exePath = [System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName
    $form.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($exePath)
    $form.ShowDialog() | Out-Null
}

Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - Listening for shortcuts, and hiding this window to the systray in three seconds" -NoNewline
for ($i = 1; $i -le 3; $i++) {
    Start-Sleep -Seconds 1
    Write-Host "." -NoNewline
}
Write-Host ""

[void] [WindowHelper]::ShowWindow($global:hWnd, 0)

Add-Type -AssemblyName System.Windows.Forms

$trayIcon = New-Object System.Windows.Forms.NotifyIcon
$exePath = [System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName
$trayIcon.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($exePath)
$trayIcon.Text = "Instrumenta Keys is running"
$trayIcon.Visible = $true

$contextMenu = New-Object System.Windows.Forms.ContextMenuStrip

$exitItem = $contextMenu.Items.Add("Exit")
$exitItem.Add_Click({
    Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - Exiting application..."
   
    $logTimer.Stop()

    $trayIcon.Dispose()
    
    [System.Windows.Forms.Application]::Exit()
    Stop-Process -Id $PID -Force 
})

$editorItem = $contextMenu.Items.Add("Shortcut Editor")
$editorItem.Add_Click({
    Start-ShortcutEditor 
})

$toggleWindowItem = $contextMenu.Items.Add("Show/Hide Window")
$toggleWindowItem.Add_Click({
    $windowState = [WindowHelper]::ShowWindow($global:hWnd, 0)
    
    if ($windowState -eq 0) {
        [WindowHelper]::ShowWindow($global:hWnd, 9)
        [WindowHelper]::SetForegroundWindow($global:hWnd)
    } else {
        [WindowHelper]::ShowWindow($global:hWnd, 0)
    }
})

$showShortcutsItem = $contextMenu.Items.Add("Available shortcuts")
$showShortcutsItem.Add_Click({
    $shortcutText = "Shortcuts and actions:`n`n"

    foreach ($shortcut in $shortcuts.Keys) {
        $friendlyName = $friendlyShortcuts[$shortcut]
        $macroName = $shortcuts[$shortcut]
        $shortcutText += "$friendlyName → $macroName`n"
    }

    [System.Windows.Forms.MessageBox]::Show($shortcutText, "Shortcut list", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
})

$trayIcon.ContextMenuStrip = $contextMenu

$keyTimestamps = @{}
$comboTimeThreshold = 200
$global:ShortcutLog = New-Object System.Collections.ArrayList

$global:TriggerShortcutEditor = New-Object PSCustomObject -Property @{ Value = $false }

$logTimer = New-Object System.Windows.Forms.Timer
$logTimer.Interval = 1000  
$logTimer.Add_Tick({  
    if ($global:TriggerShortcutEditor.Value) {
        $global:TriggerShortcutEditor.Value = $false 
        Start-ShortcutEditor
        $global:ShortcutLog.Clear() 
    } 
    if ($global:ShortcutLog.Count -gt 0) {  
        foreach ($message in $global:ShortcutLog) {
            Write-Host "$message"  
        }
        $global:ShortcutLog.Clear()  
    }
})
$logTimer.Start()


function Start-ShortcutDetection {
    param($globalKeyMap, $shortcuts, $friendlyShortcuts)

    $Runspace = [runspacefactory]::CreateRunspace()
    $Runspace.Open()
    $PowerShell = [powershell]::Create().AddScript({
        param($globalKeyMap, $shortcuts, $friendlyShortcuts, $keyTimestamps, $comboTimeThreshold, $ShortcutLog, $TriggerShortcutEditor)

        Add-Type -TypeDefinition @"
        using System;
        using System.Runtime.InteropServices;
        using System.Text;

        public class KeyboardListener {
            [DllImport("user32.dll")]
            public static extern short GetAsyncKeyState(int vKey);
        }

        public class WindowHelper {
            [DllImport("user32.dll")]
            public static extern IntPtr GetForegroundWindow();

            [DllImport("user32.dll")]
            public static extern int GetWindowThreadProcessId(IntPtr hWnd, out int lpdwProcessId);
        }
"@ -Language CSharp

        function ConnectToPowerpoint {
            try {
                return [System.Runtime.Interopservices.Marshal]::GetActiveObject("PowerPoint.Application")
            } catch {
                return $null
            }
        }

        function IsPowerPointActive {
            $hWnd = [WindowHelper]::GetForegroundWindow()
            if ($null -eq $hWnd -or $hWnd -eq [IntPtr]::Zero) { 
                return "not-active"
            }

            $activeProcessId = 0  
            [WindowHelper]::GetWindowThreadProcessId($hWnd, [ref]$activeProcessId)
            $activeProcess = Get-Process -Id $activeProcessId -ErrorAction SilentlyContinue

            if ($activeProcess.ProcessName -eq "POWERPNT") {
                return "active"
            }
            return "not-active"
        }

        $waiting = $true
        $inPresentationMode = $false      

        while ($true) {
            $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'

            $ppt = ConnectToPowerpoint
            if ($null -eq $ppt) {
                if (-not $waiting) { 
                    $ShortcutLog.Add("$timestamp - Waiting for PowerPoint instance.") | Out-Null
                    $waiting = $true
                }
                Start-Sleep -Milliseconds 5000 
                continue
            }

            if ($waiting) { 
                $ShortcutLog.Add("$timestamp - PowerPoint instance found. Accepting shortcuts when PowerPoint window is active.") | Out-Null
                $waiting = $false
            }

            $testIfActive = IsPowerPointActive
            if ($testIfActive -eq "not-active") { 
                Start-Sleep -Milliseconds 1000
                continue
            }

            if ($ppt.SlideShowWindows.Count -gt 0) {
                if (-not $inPresentationMode) {
                    $ShortcutLog.Add("$timestamp - Presentation mode detected. Shortcuts are temporarily disabled.") | Out-Null
                    $inPresentationMode = $true  
                }
                Start-Sleep -Milliseconds 1000
                continue
            } else {
                if ($inPresentationMode) {
                    $ShortcutLog.Add("$timestamp - Exited presentation mode. Shortcuts are active again.") | Out-Null
                    $inPresentationMode = $false  
                }
            }

            $pressedKeys = @()
            foreach ($key in $globalKeyMap.Keys) {  
                $keyInt = [int]$globalKeyMap[$key]  
                $state = [KeyboardListener]::GetAsyncKeyState($keyInt)

                if ($state -ne 0) {
                    $pressedKeys += "$keyInt"
                    if (-not $keyTimestamps.ContainsKey($keyInt) -or ($currentTime - $keyTimestamps[$keyInt] -gt $comboTimeThreshold)) {
                        $keyTimestamps[$keyInt] = $currentTime
                    }
                }
            }

            foreach ($virtualKeyCombo in $shortcuts.Keys) {

                $keys = $virtualKeyCombo -split ' '
                $allPressed = -not (Compare-Object -ReferenceObject $pressedKeys -DifferenceObject $keys)

                if ($allPressed) {
                    $macroName = $shortcuts[$virtualKeyCombo]
                    
                    if ($macroName -eq "InstrumentaKeysEditor") {
                        if (-not $global:TriggerShortcutEditor.Value) {  
                            $global:TriggerShortcutEditor.Value = $true
                        }

                    } else { 
                    $ShortcutLog.Add("$timestamp - Detected shortcut $($friendlyShortcuts[$virtualKeyCombo]), executing macro $macroName") | Out-Null
                    try {
                        $ppt.Run($macroName)
                    } catch {
                        $ShortcutLog.Add("$timestamp - ERROR: Failed execution of $macroName with message $_") | Out-Null
                    }

                    }

                    Start-Sleep -Milliseconds 300
                    break
                }
            }

            Start-Sleep -Milliseconds 100
        }
    })

    $PowerShell.AddArgument($globalKeyMap)
    $PowerShell.AddArgument($shortcuts)
    $PowerShell.AddArgument($friendlyShortcuts)
    $PowerShell.AddArgument($keyTimestamps)
    $PowerShell.AddArgument($comboTimeThreshold)
    $PowerShell.AddArgument($global:ShortcutLog)
    $PowerShell.AddArgument($global:TriggerShortcutEditor)
    

    $PowerShell.Runspace = $Runspace
    return $PowerShell.BeginInvoke()
}

$ShortcutDetectionHandle = Start-ShortcutDetection -globalKeyMap $global:keyMap -shortcuts $shortcuts -friendlyShortcuts $friendlyShortcuts

[System.Windows.Forms.Application]::Run()