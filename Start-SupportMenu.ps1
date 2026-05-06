param(
    [string]$ConfigPath = (Join-Path $PSScriptRoot "support-tools.json"),
    [switch]$ShowConsole
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type @"
using System;
using System.Runtime.InteropServices;

public static class Win32Window {
    [DllImport("kernel32.dll")]
    public static extern IntPtr GetConsoleWindow();

    [DllImport("user32.dll")]
    public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
}
"@

if (-not $ShowConsole) {
    $consoleHandle = [Win32Window]::GetConsoleWindow()
    if ($consoleHandle -ne [IntPtr]::Zero) {
        [Win32Window]::ShowWindow($consoleHandle, 0) | Out-Null
    }
}

function Read-ToolConfig {
    param([string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) {
        throw "Support tool config not found: $Path"
    }

    $raw = Get-Content -LiteralPath $Path -Raw | ConvertFrom-Json
    if ($null -eq $raw.tools -or $raw.tools.Count -eq 0) {
        throw "No tools were found in $Path"
    }

    return $raw
}

function Resolve-ToolPath {
    param(
        [string]$BaseDir,
        [string]$Path
    )

    if ([string]::IsNullOrWhiteSpace($Path)) {
        return $Path
    }

    if ([System.IO.Path]::IsPathRooted($Path)) {
        return $Path
    }

    return [System.IO.Path]::GetFullPath((Join-Path $BaseDir $Path))
}

function Get-OptionalPropertyValue {
    param(
        [object]$Object,
        [string]$Name,
        $DefaultValue = $null
    )

    if ($null -eq $Object) {
        return $DefaultValue
    }

    $property = $Object.PSObject.Properties[$Name]
    if ($null -eq $property) {
        return $DefaultValue
    }

    return $property.Value
}

function ConvertTo-PowerShellQuotedArgument {
    param([AllowNull()][string]$Value)

    if ($null -eq $Value) {
        return "''"
    }

    return "'" + ($Value -replace "'", "''") + "'"
}

function Convert-ToolForUi {
    param(
        [object]$Tool,
        [string]$BaseDir
    )

    $toolId = [string](Get-OptionalPropertyValue -Object $Tool -Name "id" -DefaultValue "")
    $toolName = [string](Get-OptionalPropertyValue -Object $Tool -Name "name" -DefaultValue "")
    $toolCategory = [string](Get-OptionalPropertyValue -Object $Tool -Name "category" -DefaultValue "")
    $toolDescription = [string](Get-OptionalPropertyValue -Object $Tool -Name "description" -DefaultValue "")
    $toolLaunchType = [string](Get-OptionalPropertyValue -Object $Tool -Name "launchType" -DefaultValue "")
    $toolTarget = [string](Get-OptionalPropertyValue -Object $Tool -Name "target" -DefaultValue "")
    $toolWorkingDirectory = [string](Get-OptionalPropertyValue -Object $Tool -Name "workingDirectory" -DefaultValue "")
    $toolUrl = [string](Get-OptionalPropertyValue -Object $Tool -Name "url" -DefaultValue "")
    $toolRequiresShell = [bool](Get-OptionalPropertyValue -Object $Tool -Name "requiresShell" -DefaultValue $false)
    $toolOpenInBrowser = [bool](Get-OptionalPropertyValue -Object $Tool -Name "openInBrowser" -DefaultValue $false)
    $toolKeepWindowOpen = [bool](Get-OptionalPropertyValue -Object $Tool -Name "keepWindowOpen" -DefaultValue $false)
    $toolArguments = @(Get-OptionalPropertyValue -Object $Tool -Name "arguments" -DefaultValue @())

    $resolvedTarget = Resolve-ToolPath -BaseDir $BaseDir -Path $toolTarget
    $resolvedWorkingDirectory = Resolve-ToolPath -BaseDir $BaseDir -Path $toolWorkingDirectory
    $createdUtc = $null

    if (-not [string]::IsNullOrWhiteSpace($resolvedTarget) -and (Test-Path -LiteralPath $resolvedTarget)) {
        $createdUtc = (Get-Item -LiteralPath $resolvedTarget).CreationTimeUtc
    }

    [pscustomobject]@{
        Id = $toolId
        Name = $toolName
        Category = $toolCategory
        Description = $toolDescription
        LaunchType = $toolLaunchType
        Target = $resolvedTarget
        Arguments = $toolArguments
        WorkingDirectory = $resolvedWorkingDirectory
        RequiresShell = $toolRequiresShell
        OpenInBrowser = $toolOpenInBrowser
        KeepWindowOpen = $toolKeepWindowOpen
        Url = $toolUrl
        CreatedUtc = $createdUtc
    }
}

function Get-ToolSortValue {
    param([object]$Tool)

    if ($null -eq $Tool.CreatedUtc) {
        return [datetime]::MaxValue
    }

    return [datetime]$Tool.CreatedUtc
}

function Start-SupportTool {
    param(
        [object]$Tool,
        [System.Windows.Forms.Label]$StatusLabel
    )

    try {
        $workingDirectory = $Tool.WorkingDirectory
        if ([string]::IsNullOrWhiteSpace($workingDirectory)) {
            $workingDirectory = Split-Path -Path $ConfigPath -Parent
        }

        $launchType = [string]$Tool.LaunchType
        if ($null -eq $launchType) {
            $launchType = ""
        }

        switch ($launchType.ToLowerInvariant()) {
            "powershell-script" {
                if (-not (Test-Path -LiteralPath $Tool.Target)) {
                    throw "Script not found: $($Tool.Target)"
                }

                $scriptLiteral = ConvertTo-PowerShellQuotedArgument -Value $Tool.Target
                $argumentLiterals = @($Tool.Arguments | ForEach-Object { ConvertTo-PowerShellQuotedArgument -Value ([string]$_) })
                $invokeCommand = "& $scriptLiteral"
                if ($argumentLiterals.Count -gt 0) {
                    $invokeCommand += " " + ($argumentLiterals -join " ")
                }

                $psArgs = @("-NoProfile", "-ExecutionPolicy", "Bypass")
                if ($Tool.KeepWindowOpen) {
                    $psArgs += "-NoExit"
                }
                $psArgs += @("-Command", $invokeCommand)

                Start-Process -FilePath "powershell.exe" -ArgumentList $psArgs -WorkingDirectory $workingDirectory | Out-Null
            }
            "program" {
                if ([string]::IsNullOrWhiteSpace($Tool.Target)) {
                    throw "Program target is missing."
                }

                Start-Process -FilePath $Tool.Target -ArgumentList $Tool.Arguments -WorkingDirectory $workingDirectory | Out-Null
            }
            "url" {
                if ([string]::IsNullOrWhiteSpace($Tool.Url)) {
                    throw "URL is missing."
                }

                Start-Process -FilePath $Tool.Url | Out-Null
            }
            default {
                throw "Unsupported launchType: $($Tool.LaunchType)"
            }
        }

        if ($Tool.OpenInBrowser -and -not [string]::IsNullOrWhiteSpace($Tool.Url)) {
            Start-Process -FilePath $Tool.Url | Out-Null
        }

        $StatusLabel.Text = "Launched $($Tool.Name)."
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show(
            $_.Exception.Message,
            "Launch Failed",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
        $StatusLabel.Text = "Launch failed for $($Tool.Name)."
    }
}

function New-ToolCard {
    param(
        [object]$Tool,
        [System.Windows.Forms.Label]$StatusLabel
    )

    $panel = New-Object System.Windows.Forms.Panel
    $panel.Width = 340
    $panel.Height = 170
    $panel.Margin = New-Object System.Windows.Forms.Padding(10)
    $panel.Padding = New-Object System.Windows.Forms.Padding(14)
    $panel.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#fffaf3")
    $panel.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $panel.Cursor = [System.Windows.Forms.Cursors]::Hand
    $panel.Tag = $Tool

    $category = New-Object System.Windows.Forms.Label
    $category.AutoSize = $true
    $category.Text = if ([string]::IsNullOrWhiteSpace($Tool.Category)) { "General" } else { $Tool.Category.ToUpperInvariant() }
    $category.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#1f5c49")
    $category.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
    $category.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#dfeee6")
    $category.Padding = New-Object System.Windows.Forms.Padding(8, 4, 8, 4)
    $category.Location = New-Object System.Drawing.Point(14, 14)

    $title = New-Object System.Windows.Forms.Label
    $title.AutoSize = $false
    $title.Width = 290
    $title.Height = 28
    $title.Text = $Tool.Name
    $title.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 12)
    $title.Location = New-Object System.Drawing.Point(14, 46)
    $title.Cursor = [System.Windows.Forms.Cursors]::Hand

    $description = New-Object System.Windows.Forms.Label
    $description.AutoSize = $false
    $description.Width = 300
    $description.Height = 52
    $description.Text = $Tool.Description
    $description.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#5f675f")
    $description.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $description.Location = New-Object System.Drawing.Point(14, 76)
    $description.Cursor = [System.Windows.Forms.Cursors]::Hand

    $created = New-Object System.Windows.Forms.Label
    $created.AutoSize = $false
    $created.Width = 300
    $created.Height = 20
    $created.Text = if ($null -eq $Tool.CreatedUtc) { "Created: Unknown" } else { "Created: " + ([datetime]$Tool.CreatedUtc).ToLocalTime().ToString("M/d/yyyy h:mm tt") }
    $created.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#7b746b")
    $created.Font = New-Object System.Drawing.Font("Segoe UI", 8.5)
    $created.Location = New-Object System.Drawing.Point(14, 128)
    $created.Cursor = [System.Windows.Forms.Cursors]::Hand

    $launchButton = New-Object System.Windows.Forms.Button
    $launchButton.Text = "Launch"
    $launchButton.Width = 86
    $launchButton.Height = 30
    $launchButton.Location = New-Object System.Drawing.Point(234, 124)
    $launchButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#1f5c49")
    $launchButton.ForeColor = [System.Drawing.Color]::White
    $launchButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $launchButton.Cursor = [System.Windows.Forms.Cursors]::Hand
    $launchButton.Add_Click({
        Start-SupportTool -Tool $this.Parent.Tag -StatusLabel $StatusLabel
    })

    $launchAction = {
        Start-SupportTool -Tool $this.Tag -StatusLabel $StatusLabel
    }

    $panel.Add_Click($launchAction)
    $title.Add_Click({
        Start-SupportTool -Tool $this.Parent.Tag -StatusLabel $StatusLabel
    })
    $description.Add_Click({
        Start-SupportTool -Tool $this.Parent.Tag -StatusLabel $StatusLabel
    })
    $created.Add_Click({
        Start-SupportTool -Tool $this.Parent.Tag -StatusLabel $StatusLabel
    })

    $panel.Controls.Add($category)
    $panel.Controls.Add($title)
    $panel.Controls.Add($description)
    $panel.Controls.Add($created)
    $panel.Controls.Add($launchButton)

    return $panel
}

function Render-ToolCards {
    param(
        [System.Windows.Forms.FlowLayoutPanel]$Container,
        [object[]]$Tools,
        [string]$FilterText,
        [System.Windows.Forms.Label]$StatusLabel,
        [System.Windows.Forms.Label]$CountLabel
    )

    $needle = ""
    if (-not [string]::IsNullOrWhiteSpace($FilterText)) {
        $needle = $FilterText.Trim().ToLowerInvariant()
    }

    $filteredTools = @(
        $Tools |
            Where-Object {
                if ([string]::IsNullOrWhiteSpace($needle)) {
                    return $true
                }

                $haystack = @(
                    $_.Name,
                    $_.Category,
                    $_.Description,
                    $_.Target,
                    $_.LaunchType
                ) -join " "

                return $haystack.ToLowerInvariant().Contains($needle)
            } |
            Sort-Object @{ Expression = { Get-ToolSortValue -Tool $_ } ; Ascending = $true }, @{ Expression = "Name"; Ascending = $true }
    )

    $Container.SuspendLayout()
    try {
        $Container.Controls.Clear()

        foreach ($tool in $filteredTools) {
            $Container.Controls.Add((New-ToolCard -Tool $tool -StatusLabel $StatusLabel))
        }
    }
    finally {
        $Container.ResumeLayout()
    }

    $CountLabel.Text = "{0} tool(s)" -f $filteredTools.Count

    if ($filteredTools.Count -eq 0) {
        $StatusLabel.Text = "No tools match the current filter."
    }
    elseif ($StatusLabel.Text -like "No tools match*") {
        $StatusLabel.Text = "Ready."
    }
}

$config = Read-ToolConfig -Path $ConfigPath
$configBaseDir = Split-Path -Path $ConfigPath -Parent
if ([string]::IsNullOrWhiteSpace($configBaseDir)) {
    $configBaseDir = $PSScriptRoot
}

$tools = @(
    $config.tools | ForEach-Object {
        Convert-ToolForUi -Tool $_ -BaseDir $configBaseDir
    }
)

$form = New-Object System.Windows.Forms.Form
$form.Text = "Welch Support Menu"
$form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
$form.Size = New-Object System.Drawing.Size(1160, 760)
$form.MinimumSize = New-Object System.Drawing.Size(900, 620)
$form.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#f4efe8")

$headerPanel = New-Object System.Windows.Forms.Panel
$headerPanel.Dock = [System.Windows.Forms.DockStyle]::Top
$headerPanel.Height = 118
$headerPanel.Padding = New-Object System.Windows.Forms.Padding(18, 18, 18, 12)
$headerPanel.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#f7f2ea")

$titleLabel = New-Object System.Windows.Forms.Label
$titleLabel.AutoSize = $true
$titleLabel.Text = "Welch Support Menu"
$titleLabel.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 22)
$titleLabel.Location = New-Object System.Drawing.Point(18, 14)

$subtitleLabel = New-Object System.Windows.Forms.Label
$subtitleLabel.AutoSize = $false
$subtitleLabel.Width = 760
$subtitleLabel.Height = 22
$subtitleLabel.Text = "Click any card to launch a support tool. Tools are sorted by the target file's created date."
$subtitleLabel.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#5f675f")
$subtitleLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9.5)
$subtitleLabel.Location = New-Object System.Drawing.Point(20, 54)

$filterBox = New-Object System.Windows.Forms.TextBox
$filterBox.Width = 340
$filterBox.Height = 32
$filterBox.Location = New-Object System.Drawing.Point(20, 78)
$filterBox.Font = New-Object System.Drawing.Font("Segoe UI", 10)

$reloadButton = New-Object System.Windows.Forms.Button
$reloadButton.Text = "Reload Config"
$reloadButton.Width = 120
$reloadButton.Height = 32
$reloadButton.Location = New-Object System.Drawing.Point(372, 78)
$reloadButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#a85b35")
$reloadButton.ForeColor = [System.Drawing.Color]::White
$reloadButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat

$exitButton = New-Object System.Windows.Forms.Button
$exitButton.Text = "Exit"
$exitButton.Width = 90
$exitButton.Height = 32
$exitButton.Location = New-Object System.Drawing.Point(500, 78)
$exitButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#5f675f")
$exitButton.ForeColor = [System.Drawing.Color]::White
$exitButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat

$countLabel = New-Object System.Windows.Forms.Label
$countLabel.AutoSize = $true
$countLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$countLabel.Location = New-Object System.Drawing.Point(602, 83)

$headerPanel.Controls.Add($titleLabel)
$headerPanel.Controls.Add($subtitleLabel)
$headerPanel.Controls.Add($filterBox)
$headerPanel.Controls.Add($reloadButton)
$headerPanel.Controls.Add($exitButton)
$headerPanel.Controls.Add($countLabel)

$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.Dock = [System.Windows.Forms.DockStyle]::Bottom
$statusLabel.Height = 32
$statusLabel.Padding = New-Object System.Windows.Forms.Padding(18, 8, 18, 0)
$statusLabel.Text = "Ready."
$statusLabel.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#5f675f")
$statusLabel.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#f7f2ea")

$scrollPanel = New-Object System.Windows.Forms.FlowLayoutPanel
$scrollPanel.Dock = [System.Windows.Forms.DockStyle]::Fill
$scrollPanel.AutoScroll = $true
$scrollPanel.WrapContents = $true
$scrollPanel.Padding = New-Object System.Windows.Forms.Padding(12)
$scrollPanel.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#f4efe8")

$reloadAction = {
    try {
        $script:config = Read-ToolConfig -Path $ConfigPath
        $script:tools = @(
            $script:config.tools | ForEach-Object {
                Convert-ToolForUi -Tool $_ -BaseDir $configBaseDir
            }
        )
        Render-ToolCards -Container $scrollPanel -Tools $script:tools -FilterText $filterBox.Text -StatusLabel $statusLabel -CountLabel $countLabel
        $statusLabel.Text = "Config reloaded."
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show(
            $_.Exception.Message,
            "Reload Failed",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
        $statusLabel.Text = "Config reload failed."
    }
}

$filterBox.Add_TextChanged({
    Render-ToolCards -Container $scrollPanel -Tools $script:tools -FilterText $filterBox.Text -StatusLabel $statusLabel -CountLabel $countLabel
})

$reloadButton.Add_Click($reloadAction)
$exitButton.Add_Click({
    $form.Close()
})

$form.Controls.Add($scrollPanel)
$form.Controls.Add($statusLabel)
$form.Controls.Add($headerPanel)

Render-ToolCards -Container $scrollPanel -Tools $tools -FilterText "" -StatusLabel $statusLabel -CountLabel $countLabel

[void]$form.ShowDialog()
