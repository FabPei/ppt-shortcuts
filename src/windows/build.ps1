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

    // FIX: Added VkKeyScan to translate a character to a virtual-key code.
    // This respects the user's current keyboard layout (e.g., QWERTZ vs QWERTY)
    // and resolves issues with keys like 'Y' and 'Z' on German keyboards.
    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    public static extern short VkKeyScan(char ch);
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

$instrumentaKeysVersion = "0.21"

Write-Host "██╗███╗   ██╗███████╗████████╗██████╗ ██╗   ██╗███╗   ███╗███████╗███╗   ██╗████████╗ █████╗ "
Write-Host "██║████╗  ██║██╔════╝╚══██╔══╝██╔══██╗██║   ██║████╗ ████║██╔════╝████╗  ██║╚══██╔══╝██╔══██╗"
Write-Host "██║██╔██╗ ██║███████╗   ██║    ██████╔╝██║   ██║██╔████╔██║█████╗  ██╔██╗ ██║   ██║    ███████║"
Write-Host "██║██║╚██╗██║╚════██║   ██║    ██╔══██╗██║   ██║██║╚██╔╝██║██╔══╝  ██║╚██╗██║   ██║    ██╔══██║"
Write-Host "██║██║ ╚████║███████║   ██║    ██║  ██║╚██████╔╝██║ ╚═╝ ██║███████╗██║ ╚████║   ██║    ██║  ██║"
Write-Host "╚═╝╚═╝  ╚═══╝╚══════╝   ╚═╝    ╚═╝  ╚═╝ ╚═════╝ ╚═╝     ╚═╝╚══════╝╚═╝  ╚═══╝   ╚═╝    ╚═╝  ╚═╝"
Write-Host "██╗  ██╗███████╗██╗    ██╗███████╗                                                           "
Write-Host "██║ ██╔╝██╔════╝╚██╗ ██╔╝██╔════╝                                                           "
Write-Host "█████╔╝ █████╗   ╚████╔╝ ███████╗      Keyboard Shortcut Companion (v. $instrumentaKeysVersion)"
Write-Host "██╔═██╗ ██╔══╝    ╚██╔╝  ╚════██║                                                           "
Write-Host "██║  ██╗███████╗   ██║   ███████║                                                           "
Write-Host "╚═╝  ╚═╝╚══════╝   ╚═╝   ╚══════╝                                                           "
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
                    $trimmedKey = $key.Trim()
                    if (-not $global:keyMap.ContainsKey($trimmedKey)) {
                        try {
                            # Map special keys by name, and all other character keys dynamically
                            switch ($trimmedKey.ToUpper()) {
                                "CTRL"           { $global:keyMap[$trimmedKey] = 0x11 }
                                "SHIFT"          { $global:keyMap[$trimmedKey] = 0x10 }
                                "ALT"            { $global:keyMap[$trimmedKey] = 0x12 }
                                "DEL"            { $global:keyMap[$trimmedKey] = 0x2E }
                                "UP"             { $global:keyMap[$trimmedKey] = 0x26 }
                                "DOWN"           { $global:keyMap[$trimmedKey] = 0x28 }
                                "LEFT"           { $global:keyMap[$trimmedKey] = 0x25 }
                                "RIGHT"          { $global:keyMap[$trimmedKey] = 0x27 }
                                "ESC"            { $global:keyMap[$trimmedKey] = 0x1B }
                                "ENTER"          { $global:keyMap[$trimmedKey] = 0x0D }
                                "TAB"            { $global:keyMap[$trimmedKey] = 0x09 }
                                "SPACE"          { $global:keyMap[$trimmedKey] = 0x20 }
                                "BACKSPACE"      { $global:keyMap[$trimmedKey] = 0x08 }
                                "PAGEUP"         { $global:keyMap[$trimmedKey] = 0x21 }
                                "PAGEDOWN"       { $global:keyMap[$trimmedKey] = 0x22 }
                                "HOME"           { $global:keyMap[$trimmedKey] = 0x24 }
                                "END"            { $global:keyMap[$trimmedKey] = 0x23 }
                                "INSERT"         { $global:keyMap[$trimmedKey] = 0x2D }
                                "F1"             { $global:keyMap[$trimmedKey] = 0x70 }
                                "F2"             { $global:keyMap[$trimmedKey] = 0x71 }
                                "F3"             { $global:keyMap[$trimmedKey] = 0x72 }
                                "F4"             { $global:keyMap[$trimmedKey] = 0x73 }
                                "F5"             { $global:keyMap[$trimmedKey] = 0x74 }
                                "F6"             { $global:keyMap[$trimmedKey] = 0x75 }
                                "F7"             { $global:keyMap[$trimmedKey] = 0x76 }
                                "F8"             { $global:keyMap[$trimmedKey] = 0x77 }
                                "F9"             { $global:keyMap[$trimmedKey] = 0x78 }
                                "F10"            { $global:keyMap[$trimmedKey] = 0x79 }
                                "F11"            { $global:keyMap[$trimmedKey] = 0x7A }
                                "F12"            { $global:keyMap[$trimmedKey] = 0x7B }
                                "NUM0"           { $global:keyMap[$trimmedKey] = 0x60 }
                                "NUM1"           { $global:keyMap[$trimmedKey] = 0x61 }
                                "NUM2"           { $global:keyMap[$trimmedKey] = 0x62 }
                                "NUM3"           { $global:keyMap[$trimmedKey] = 0x63 }
                                "NUM4"           { $global:keyMap[$trimmedKey] = 0x64 }
                                "NUM5"           { $global:keyMap[$trimmedKey] = 0x65 }
                                "NUM6"           { $global:keyMap[$trimmedKey] = 0x66 }
                                "NUM7"           { $global:keyMap[$trimmedKey] = 0x67 }
                                "NUM8"           { $global:keyMap[$trimmedKey] = 0x68 }
                                "NUM9"           { $global:keyMap[$trimmedKey] = 0x69 }
                                "NUMLOCK"        { $global:keyMap[$trimmedKey] = 0x90 }
                                "NUMPADDIVIDE"   { $global:keyMap[$trimmedKey] = 0x6F }
                                "NUMPADMULTIPLY" { $global:keyMap[$trimmedKey] = 0x6A }
                                "NUMPADSUBTRACT" { $global:keyMap[$trimmedKey] = 0x6D }
                                "NUMPADADD"      { $global:keyMap[$trimmedKey] = 0x6B }
                                "NUMPADENTER"    { $global:keyMap[$trimmedKey] = 0x0D }
                                "CAPSLOCK"       { $global:keyMap[$trimmedKey] = 0x14 }
                                "SCROLLLOCK"     { $global:keyMap[$trimmedKey] = 0x91 }
                                "PRINTSCREEN"    { $global:keyMap[$trimmedKey] = 0x2C }
                                "PAUSEBREAK"     { $global:keyMap[$trimmedKey] = 0x13 }
                                Default {
                                    # FIX: Dynamically map character keys (A-Z, 0-9, symbols) using the current keyboard layout.
                                    # This replaces the hardcoded list and fixes the Y/Z issue on German keyboards.
                                    if ($trimmedKey.Length -eq 1) {
                                        $char = $trimmedKey[0]
                                        $vkScanResult = [KeyboardListener]::VkKeyScan($char)
                                        if ($vkScanResult -ne -1) {
                                            # The virtual key code is in the low byte of the result
                                            $vkCode = $vkScanResult -band 0xFF
                                            $global:keyMap[$trimmedKey] = $vkCode
                                        } else {
                                            Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - WARNING: Could not map character '$trimmedKey' to a virtual key."
                                        }
                                    } else {
                                        Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - ERROR: Ignoring unsupported key '$trimmedKey'."
                                    }
                                }
                            }
                        } catch {
                            Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - ERROR: Failed to process key '$trimmedKey' in shortcut '$($entry.Key)'."
                        }
                    }

                    if ($global:keyMap.ContainsKey($trimmedKey)) {
                        $virtualKeys += $global:keyMap[$trimmedKey]
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
    $repoUrl = "https://api.github.com/repos/iappyx/Instrumenta-Keys/contents/shared-shortcuts/windows/"

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

function Export-PowerPointMacros {
    Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - Attempting to list all available PowerPoint macros..."
    
    try {
        $ppt = [System.Runtime.Interopservices.Marshal]::GetActiveObject("PowerPoint.Application")
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Could not connect to a running PowerPoint instance. Please make sure PowerPoint is open.", "Connection Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    $macroList = New-Object System.Collections.ArrayList

    # Constant for vbext_ProcKind `vbext_pk_Proc` (Sub or Function)
    $vbext_pk_Proc = 1 

    # --- Gather Macros from Add-Ins ---
    Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - Searching for macros in Add-Ins..."
    foreach ($addin in $ppt.AddIns) {
        if ($addin.Loaded) {
            try {
                if ($addin.Object.HasVBProject) {
                    $vbProject = $addin.Object.VBProject
                    foreach ($component in $vbProject.VBComponents) {
                        $codeModule = $component.CodeModule
                        $lineNum = 1
                        while ($lineNum -lt $codeModule.CountOfLines) {
                            $procName = $codeModule.ProcOfLine($lineNum, [ref]$vbext_pk_Proc)
                            if ($null -ne $procName) {
                                $procBodyLineCount = $codeModule.ProcCountLines($procName, [ref]$vbext_pk_Proc)
                                $macroList.Add([PSCustomObject]@{
                                    Source     = $addin.Name
                                    Module     = $component.Name
                                    MacroName  = $procName
                                }) | Out-Null
                                $lineNum += $procBodyLineCount
                            } else {
                                $lineNum++
                            }
                        }
                    }
                }
            } catch {
                Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - Warning: Could not inspect Add-In '$($addin.Name)'. It might be protected."
            }
        }
    }

    # --- Gather Macros from Open Presentations ---
     Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - Searching for macros in open presentations..."
    foreach ($presentation in $ppt.Presentations) {
        try {
            if ($presentation.HasVBProject) {
                $vbProject = $presentation.VBProject
                foreach ($component in $vbProject.VBComponents) {
                    $codeModule = $component.CodeModule
                    $lineNum = 1
                    while ($lineNum -lt $codeModule.CountOfLines) {
                        $procName = $codeModule.ProcOfLine($lineNum, [ref]$vbext_pk_Proc)
                        if ($null -ne $procName) {
                            $procBodyLineCount = $codeModule.ProcCountLines($procName, [ref]$vbext_pk_Proc)
                            $macroList.Add([PSCustomObject]@{
                                Source     = $presentation.Name
                                Module     = $component.Name
                                MacroName  = $procName
                            }) | Out-Null
                            $lineNum += $procBodyLineCount
                        } else {
                            $lineNum++
                        }
                    }
                }
            }
        } catch {
             Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - Warning: Could not inspect presentation '$($presentation.Name)'. It might be protected."
        }
    }

    if ($macroList.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No macros were found in loaded add-ins or open presentations.", "No Macros Found", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }

    # --- Display Form ---
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Available PowerPoint Macros"
    $form.Width = 800
    $form.Height = 500
    $form.StartPosition = "CenterScreen"
    $form.TopMost = $true

    $grid = New-Object System.Windows.Forms.DataGridView
    $grid.Dock = "Fill"
    $grid.DataSource = $macroList
    $grid.AutoSizeColumnsMode = "Fill"
    $grid.AllowUserToAddRows = $false
    $grid.ReadOnly = $true

    $exportButton = New-Object System.Windows.Forms.Button
    $exportButton.Text = "Export to CSV"
    $exportButton.Dock = "Bottom"
    $exportButton.Height = 30
    $exportButton.Add_Click({
        $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
        $saveFileDialog.Filter = "CSV file (*.csv)|*.csv"
        $saveFileDialog.Title = "Export Macros As"
        if ($saveFileDialog.ShowDialog() -eq "OK") {
            $macroList | Export-Csv -Path $saveFileDialog.FileName -NoTypeInformation
            [System.Windows.Forms.MessageBox]::Show("Macro list exported successfully!", "Export Complete")
        }
    })

    $form.Controls.Add($grid)
    $form.Controls.Add($exportButton)
    $exePath = [System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName
    $form.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($exePath)
    $form.ShowDialog()
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

$listMacrosItem = $contextMenu.Items.Add("List PowerPoint Macros")
$listMacrosItem.Add_Click({
    Export-PowerPointMacros
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
            $currentTime = (Get-Date).Ticks / [System.TimeSpan]::TicksPerMillisecond
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
                # Compare-Object returns null if collections are identical
                $allPressed = $null -eq (Compare-Object -ReferenceObject $pressedKeys -DifferenceObject $keys -PassThru | Where-Object { $_ -in $keys })

                if ($allPressed -and $pressedKeys.Count -eq $keys.Count) {
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

