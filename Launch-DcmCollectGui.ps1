[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

function Get-ScriptDirectory {
    if ($PSScriptRoot -and (Test-Path -LiteralPath $PSScriptRoot -PathType Container)) {
        return $PSScriptRoot
    }

    $myCommandPath = $null
    if ($MyInvocation -and $MyInvocation.MyCommand) {
        if ($MyInvocation.MyCommand -is [string]) {
            $myCommandPath = $MyInvocation.MyCommand
        }
        elseif ($MyInvocation.MyCommand.PSObject.Properties.Match("Path").Count -gt 0) {
            $myCommandPath = $MyInvocation.MyCommand.Path
        }
    }

    if (-not [string]::IsNullOrWhiteSpace($myCommandPath)) {
        return (Split-Path -Parent $myCommandPath)
    }

    # Fallback for compiled executable context.
    return [System.AppContext]::BaseDirectory.TrimEnd([char]'\')
}

$scriptDir = Get-ScriptDirectory
$backendExe = Join-Path $scriptDir "dcmcollect.exe"
$backendPs1 = Join-Path $scriptDir "Collect-DicomMedia.ps1"
$iconPath = Join-Path $scriptDir "assets\dcmcollect.ico"
$script:diagLogPath = Join-Path $scriptDir "gui-crash.log"
$script:lastDest = $null

function Write-Diagnostic {
    param([string]$Message)

    try {
        $line = "{0} {1}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss.fff"), $Message
        Add-Content -LiteralPath $script:diagLogPath -Encoding UTF8 -Value $line
    }
    catch {
        # Never throw from diagnostics.
    }
}

Write-Diagnostic "GUI started"

function ConvertTo-QuotedArgument {
    param([Parameter(Mandatory = $true)][string]$Value)

    if ($Value -notmatch '[\s"]') {
        return $Value
    }

    $escaped = $Value -replace '"', '\"'
    return '"{0}"' -f $escaped
}

$form = New-Object System.Windows.Forms.Form
$form.Text = "dcmcollect"
$form.StartPosition = "CenterScreen"
$form.Size = New-Object System.Drawing.Size(820, 620)
$form.MinimumSize = New-Object System.Drawing.Size(820, 620)

if (Test-Path -LiteralPath $iconPath -PathType Leaf) {
    try {
        $form.Icon = New-Object System.Drawing.Icon($iconPath)
    }
    catch {
        # Ignore icon load failures and keep default icon.
    }
}

$font = New-Object System.Drawing.Font("Segoe UI", 9)
$form.Font = $font

$lblIntro = New-Object System.Windows.Forms.Label
$lblIntro.Text = "This tool collects DICOM files into a flat media folder, builds a DICOMDIR index, and can bundle Weasis for one-click viewing. Select source and destination, choose options, then run."
$lblIntro.Location = New-Object System.Drawing.Point(16, 12)
$lblIntro.Size = New-Object System.Drawing.Size(779, 42)

$lblSrc = New-Object System.Windows.Forms.Label
$lblSrc.Text = "Source Folder"
$lblSrc.Location = New-Object System.Drawing.Point(16, 58)
$lblSrc.Size = New-Object System.Drawing.Size(120, 22)

$txtSrc = New-Object System.Windows.Forms.TextBox
$txtSrc.Location = New-Object System.Drawing.Point(16, 80)
$txtSrc.Size = New-Object System.Drawing.Size(680, 26)

$btnSrc = New-Object System.Windows.Forms.Button
$btnSrc.Text = "Browse..."
$btnSrc.Location = New-Object System.Drawing.Point(705, 78)
$btnSrc.Size = New-Object System.Drawing.Size(90, 28)

$lblDest = New-Object System.Windows.Forms.Label
$lblDest.Text = "Destination Folder"
$lblDest.Location = New-Object System.Drawing.Point(16, 116)
$lblDest.Size = New-Object System.Drawing.Size(140, 22)

$txtDest = New-Object System.Windows.Forms.TextBox
$txtDest.Location = New-Object System.Drawing.Point(16, 138)
$txtDest.Size = New-Object System.Drawing.Size(680, 26)

$btnDest = New-Object System.Windows.Forms.Button
$btnDest.Text = "Browse..."
$btnDest.Location = New-Object System.Drawing.Point(705, 136)
$btnDest.Size = New-Object System.Drawing.Size(90, 28)

$chkPackage = New-Object System.Windows.Forms.CheckBox
$chkPackage.Text = "Package Weasis"
$chkPackage.Location = New-Object System.Drawing.Point(16, 174)
$chkPackage.Size = New-Object System.Drawing.Size(130, 24)

$chkVerify = New-Object System.Windows.Forms.CheckBox
$chkVerify.Text = "Verify DICOMDIR"
$chkVerify.Location = New-Object System.Drawing.Point(156, 174)
$chkVerify.Size = New-Object System.Drawing.Size(140, 24)

$chkWeasisOnly = New-Object System.Windows.Forms.CheckBox
$chkWeasisOnly.Text = "Weasis Only"
$chkWeasisOnly.Location = New-Object System.Drawing.Point(306, 174)
$chkWeasisOnly.Size = New-Object System.Drawing.Size(110, 24)

$chkDeidentify = New-Object System.Windows.Forms.CheckBox
$chkDeidentify.Text = "De-identify Data"
$chkDeidentify.Location = New-Object System.Drawing.Point(426, 174)
$chkDeidentify.Size = New-Object System.Drawing.Size(130, 24)

$btnRun = New-Object System.Windows.Forms.Button
$btnRun.Text = "Run"
$btnRun.Location = New-Object System.Drawing.Point(16, 212)
$btnRun.Size = New-Object System.Drawing.Size(110, 34)

$btnClose = New-Object System.Windows.Forms.Button
$btnClose.Text = "Close"
$btnClose.Location = New-Object System.Drawing.Point(136, 212)
$btnClose.Size = New-Object System.Drawing.Size(110, 34)

$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(270, 212)
$progressBar.Size = New-Object System.Drawing.Size(525, 22)
$progressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Continuous

$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Text = "Ready"
$lblStatus.Location = New-Object System.Drawing.Point(270, 237)
$lblStatus.Size = New-Object System.Drawing.Size(430, 24)

$lblEta = New-Object System.Windows.Forms.Label
$lblEta.Text = "ETA: --"
$lblEta.Location = New-Object System.Drawing.Point(705, 237)
$lblEta.Size = New-Object System.Drawing.Size(90, 24)

$grpSummary = New-Object System.Windows.Forms.GroupBox
$grpSummary.Text = "Run Summary"
$grpSummary.Location = New-Object System.Drawing.Point(16, 264)
$grpSummary.Size = New-Object System.Drawing.Size(779, 64)

$lblSumCandidates = New-Object System.Windows.Forms.Label
$lblSumCandidates.Text = "Candidates: 0/0"
$lblSumCandidates.Location = New-Object System.Drawing.Point(12, 28)
$lblSumCandidates.Size = New-Object System.Drawing.Size(150, 22)

$lblSumCopied = New-Object System.Windows.Forms.Label
$lblSumCopied.Text = "Copied: 0"
$lblSumCopied.Location = New-Object System.Drawing.Point(170, 28)
$lblSumCopied.Size = New-Object System.Drawing.Size(100, 22)

$lblSumElapsed = New-Object System.Windows.Forms.Label
$lblSumElapsed.Text = "Elapsed: 00:00"
$lblSumElapsed.Location = New-Object System.Drawing.Point(278, 28)
$lblSumElapsed.Size = New-Object System.Drawing.Size(120, 22)

$lblSumRate = New-Object System.Windows.Forms.Label
$lblSumRate.Text = "Rate: 0.0 files/s"
$lblSumRate.Location = New-Object System.Drawing.Point(406, 28)
$lblSumRate.Size = New-Object System.Drawing.Size(140, 22)

$lblSumState = New-Object System.Windows.Forms.Label
$lblSumState.Text = "State: Idle"
$lblSumState.Location = New-Object System.Drawing.Point(554, 28)
$lblSumState.Size = New-Object System.Drawing.Size(200, 22)

$grpSummary.Controls.AddRange(@(
    $lblSumCandidates, $lblSumCopied, $lblSumElapsed, $lblSumRate, $lblSumState
))

$txtLog = New-Object System.Windows.Forms.TextBox
$txtLog.Location = New-Object System.Drawing.Point(16, 336)
$txtLog.Size = New-Object System.Drawing.Size(779, 235)
$txtLog.Multiline = $true
$txtLog.ScrollBars = "Vertical"
$txtLog.ReadOnly = $true
$txtLog.WordWrap = $false

$form.Controls.AddRange(@(
    $lblIntro,
    $lblSrc, $txtSrc, $btnSrc,
    $lblDest, $txtDest, $btnDest,
    $chkPackage, $chkVerify, $chkWeasisOnly, $chkDeidentify,
    $btnRun, $btnClose, $progressBar, $lblStatus, $lblEta, $grpSummary, $txtLog
))

$folderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
$uiTimer = New-Object System.Windows.Forms.Timer
$uiTimer.Interval = 1000
$uiTimer.Add_Tick({
    try {
        Drain-ProcessOutputFiles

        if ($script:currentProcess) {
            if (-not $script:currentProcess.HasExited) {
                Update-Summary -UpdateElapsed
            }
            elseif (-not $script:completionHandled) {
                Complete-Run -ExitCode $script:currentProcess.ExitCode
            }
        }
    }
    catch {
        Write-Diagnostic ("Timer tick error: " + $_.Exception.ToString())
    }
})

$script:currentProcess = $null
$script:completionHandled = $true
$script:stdoutPath = $null
$script:stderrPath = $null
$script:stdoutLineCount = 0
$script:stderrLineCount = 0
$script:progressTotal = 0
$script:progressProcessed = 0
$script:progressCopied = 0
$script:progressDiscovered = 0
$script:runStart = $null
$script:discoverySeconds = $null
$script:scanSeconds = $null
$script:dicomdirSeconds = $null

function Format-Duration {
    param([TimeSpan]$Span)

    if ($Span.TotalHours -ge 1) {
        return "{0:00}:{1:00}:{2:00}" -f [math]::Floor($Span.TotalHours), $Span.Minutes, $Span.Seconds
    }
    return "{0:00}:{1:00}" -f $Span.Minutes, $Span.Seconds
}

function Update-Summary {
    param(
        [string]$State = "",
        [switch]$UpdateElapsed
    )

    if ($script:progressTotal -gt 0) {
        $lblSumCandidates.Text = "Candidates: $($script:progressProcessed)/$($script:progressTotal)"
    }
    else {
        if ($script:progressDiscovered -gt 0) {
            $lblSumCandidates.Text = "Candidates: $($script:progressDiscovered) found"
        }
        else {
            $lblSumCandidates.Text = "Candidates: $($script:progressProcessed)/0"
        }
    }

    $lblSumCopied.Text = "Copied: $($script:progressCopied)"

    if ($script:runStart -and $UpdateElapsed) {
        $elapsed = (Get-Date) - $script:runStart
        $lblSumElapsed.Text = "Elapsed: $(Format-Duration $elapsed)"

        if ($elapsed.TotalSeconds -gt 0) {
            $rate = $script:progressCopied / $elapsed.TotalSeconds
            $lblSumRate.Text = ("Rate: {0:N2} files/s" -f $rate)
        }
    }

    if (-not [string]::IsNullOrWhiteSpace($State)) {
        $lblSumState.Text = "State: $State"
    }
}

function Update-ProgressFromLine {
    param([string]$Line)

    if ([string]::IsNullOrWhiteSpace($Line)) {
        return $false
    }

    if ($Line -match '^STAT\|TOTAL_CANDIDATES=(\d+)$') {
        $script:progressTotal = [int]$matches[1]
        $script:progressProcessed = 0
        $script:progressCopied = 0

        $progressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Continuous
        $progressBar.Minimum = 0
        $progressBar.Maximum = [Math]::Max(1, $script:progressTotal)
        $progressBar.Value = 0
        $lblStatus.Text = "Preparing scan: 0/$($script:progressTotal)"
        $lblEta.Text = "ETA: --"
        Update-Summary -State "Scanning" -UpdateElapsed
        return $true
    }

    if ($Line -match '^DISCOVER\|COUNT=(\d+)\|FILE=(.*)$') {
        $script:progressDiscovered = [int]$matches[1]
        $lblStatus.Text = "Discovering files: $($script:progressDiscovered) found"
        $lblEta.Text = "ETA: estimating..."

        if ($progressBar.Style -ne [System.Windows.Forms.ProgressBarStyle]::Marquee) {
            $progressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Marquee
            $progressBar.MarqueeAnimationSpeed = 25
        }

        Update-Summary -State "Discovering" -UpdateElapsed
        return $true
    }

    if ($Line -match '^PROGRESS\|PROCESSED=(\d+)\|COPIED=(\d+)\|CURRENT=(.*)$') {
        $script:progressProcessed = [int]$matches[1]
        $script:progressCopied = [int]$matches[2]

        if ($progressBar.Style -ne [System.Windows.Forms.ProgressBarStyle]::Continuous) {
            $progressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Continuous
        }

        if ($script:progressTotal -gt 0) {
            $value = [Math]::Min($script:progressProcessed, $progressBar.Maximum)
            $progressBar.Value = [Math]::Max(0, $value)

            $lblStatus.Text = "Scanning: $($script:progressProcessed)/$($script:progressTotal) candidates, copied: $($script:progressCopied)"

            if ($script:runStart -and $script:progressProcessed -gt 0) {
                $elapsed = (Get-Date) - $script:runStart
                $rate = $script:progressProcessed / [Math]::Max(1.0, $elapsed.TotalSeconds)
                if ($rate -gt 0) {
                    $remainingCount = [Math]::Max(0, $script:progressTotal - $script:progressProcessed)
                    $remainingSeconds = [Math]::Ceiling($remainingCount / $rate)
                    $eta = New-TimeSpan -Seconds $remainingSeconds
                    $lblEta.Text = "ETA: $(Format-Duration $eta)"
                }
            }
        }
        else {
            $lblStatus.Text = "Scanning: $($script:progressProcessed) candidates, copied: $($script:progressCopied)"
        }

        Update-Summary -State "Scanning" -UpdateElapsed
        return $true
    }

    if ($Line -match '^PHASE\|DISCOVERY_SECONDS=([0-9]+(?:\.[0-9]+)?)$') {
        $script:discoverySeconds = [double]$matches[1]
        $lblStatus.Text = ("Discovery completed in {0:N2}s" -f $script:discoverySeconds)
        Update-Summary -State "Scanning" -UpdateElapsed
        return $true
    }

    if ($Line -match '^PHASE\|SCAN_SECONDS=([0-9]+(?:\.[0-9]+)?)$') {
        $script:scanSeconds = [double]$matches[1]
        $lblStatus.Text = ("Candidate scan completed in {0:N2}s" -f $script:scanSeconds)
        Update-Summary -State "Building DICOMDIR" -UpdateElapsed
        return $true
    }

    if ($Line -match '^PHASE\|DICOMDIR_SECONDS=([0-9]+(?:\.[0-9]+)?)$') {
        $script:dicomdirSeconds = [double]$matches[1]
        $lblStatus.Text = ("DICOMDIR build completed in {0:N2}s" -f $script:dicomdirSeconds)
        Update-Summary -State "Completing" -UpdateElapsed
        return $true
    }

    return $false
}

function Add-LogLine {
    param([string]$Message)

    if ([string]::IsNullOrWhiteSpace($Message)) {
        return
    }

    if ($txtLog.InvokeRequired) {
        $null = $txtLog.BeginInvoke([Action[string]]{ param($m) Add-LogLine -Message $m }, $Message)
        return
    }

    try {
        $isProgressLine = Update-ProgressFromLine -Line $Message
        if ($isProgressLine) {
            return
        }

        $txtLog.AppendText($Message + [Environment]::NewLine)
    }
    catch {
        Write-Diagnostic ("Add-LogLine error: " + $_.Exception.ToString())
    }
}

function Drain-ProcessOutputFiles {
    if ($script:stdoutPath -and (Test-Path -LiteralPath $script:stdoutPath -PathType Leaf)) {
        try {
            $outLines = @(Get-Content -LiteralPath $script:stdoutPath -ErrorAction SilentlyContinue)
            if ($outLines.Count -gt 0) {
                for ($idx = $script:stdoutLineCount; $idx -lt $outLines.Count; $idx++) {
                    Add-LogLine ([string]$outLines[$idx])
                }
                $script:stdoutLineCount = $outLines.Count
            }
        }
        catch {
            Write-Diagnostic ("Drain stdout error: " + $_.Exception.ToString())
        }
    }

    if ($script:stderrPath -and (Test-Path -LiteralPath $script:stderrPath -PathType Leaf)) {
        try {
            $errLines = @(Get-Content -LiteralPath $script:stderrPath -ErrorAction SilentlyContinue)
            if ($errLines.Count -gt 0) {
                for ($idx = $script:stderrLineCount; $idx -lt $errLines.Count; $idx++) {
                    Add-LogLine ("[ERR] " + [string]$errLines[$idx])
                }
                $script:stderrLineCount = $errLines.Count
            }
        }
        catch {
            Write-Diagnostic ("Drain stderr error: " + $_.Exception.ToString())
        }
    }
}

function Cleanup-ProcessOutputFiles {
    foreach ($path in @($script:stdoutPath, $script:stderrPath)) {
        if ($path) {
            try {
                Remove-Item -LiteralPath $path -Force -ErrorAction SilentlyContinue
            }
            catch {
                Write-Diagnostic ("Cleanup output file error: " + $_.Exception.ToString())
            }
        }
    }

    $script:stdoutPath = $null
    $script:stderrPath = $null
    $script:stdoutLineCount = 0
    $script:stderrLineCount = 0
}

function Complete-Run {
    param([int]$ExitCode)

    if ($script:completionHandled) {
        return
    }

    $script:completionHandled = $true
    try {
        Drain-ProcessOutputFiles
        Set-RunningState -Running $false
        $lblStatus.Text = "Completed (ExitCode=$ExitCode)"

        if ($script:runStart) {
            $elapsed = (Get-Date) - $script:runStart
            Add-LogLine ("Elapsed: $(Format-Duration $elapsed)")
        }
        if ($script:discoverySeconds -ne $null) {
            Add-LogLine ("Phase timing: discovery={0:N2}s" -f $script:discoverySeconds)
        }
        if ($script:scanSeconds -ne $null) {
            Add-LogLine ("Phase timing: scan={0:N2}s" -f $script:scanSeconds)
        }
        if ($script:dicomdirSeconds -ne $null) {
            Add-LogLine ("Phase timing: dicomdir={0:N2}s" -f $script:dicomdirSeconds)
        }

        Update-Summary -State "Completed" -UpdateElapsed
        Add-LogLine ("Process exited with code $ExitCode")
        Write-Diagnostic ("Run completed with exit code " + $ExitCode)

        if ($ExitCode -eq 0 -and -not [string]::IsNullOrWhiteSpace($script:lastDest) -and (Test-Path -LiteralPath $script:lastDest -PathType Container)) {
            $choice = [System.Windows.Forms.MessageBox]::Show(
                "Processing completed successfully. Open the destination folder?",
                "dcmcollect",
                [System.Windows.Forms.MessageBoxButtons]::YesNo,
                [System.Windows.Forms.MessageBoxIcon]::Information
            )

            if ($choice -eq [System.Windows.Forms.DialogResult]::Yes) {
                Start-Process -FilePath "explorer.exe" -ArgumentList @($script:lastDest) | Out-Null
            }
        }

        $script:currentProcess = $null
        Cleanup-ProcessOutputFiles
    }
    catch {
        Write-Diagnostic ("Complete-Run error: " + $_.Exception.ToString())
    }
}

function Set-RunningState {
    param([bool]$Running)

    $btnRun.Enabled = -not $Running
    $btnSrc.Enabled = -not $Running
    $btnDest.Enabled = -not $Running
    $chkPackage.Enabled = -not $Running
    $chkVerify.Enabled = -not $Running
    $chkWeasisOnly.Enabled = -not $Running
    $chkDeidentify.Enabled = -not $Running
    $txtSrc.Enabled = -not $Running
    $txtDest.Enabled = -not $Running
    if ($Running) {
        $uiTimer.Start()
        $progressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Marquee
        $progressBar.MarqueeAnimationSpeed = 25
        $lblStatus.Text = "Starting..."
        $lblEta.Text = "ETA: --"
        $lblSumElapsed.Text = "Elapsed: 00:00"
        $lblSumRate.Text = "Rate: 0.00 files/s"
        Update-Summary -State "Running"
    }
    else {
        $uiTimer.Stop()
        $progressBar.MarqueeAnimationSpeed = 0
        $lblEta.Text = "ETA: --"
        if ($script:progressTotal -gt 0 -and $script:progressProcessed -ge $script:progressTotal) {
            $progressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Continuous
            $progressBar.Minimum = 0
            $progressBar.Maximum = $script:progressTotal
            $progressBar.Value = $script:progressTotal
        }
        else {
            $progressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Continuous
            $progressBar.Minimum = 0
            $progressBar.Maximum = 1
            $progressBar.Value = 0
        }
        if ($lblStatus.Text -eq "Starting..." -or $lblStatus.Text -eq "Running...") {
            $lblStatus.Text = "Ready"
        }
        Update-Summary -State "Idle" -UpdateElapsed
    }
}

$btnSrc.Add_Click({
    if ($folderDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtSrc.Text = $folderDialog.SelectedPath
    }
})

$btnDest.Add_Click({
    if ($folderDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtDest.Text = $folderDialog.SelectedPath
    }
})

$chkWeasisOnly.Add_CheckedChanged({
    $isWeasisOnly = $chkWeasisOnly.Checked
    $txtSrc.Enabled = -not $isWeasisOnly
    $btnSrc.Enabled = -not $isWeasisOnly
})

$btnRun.Add_Click({
    try {
        Write-Diagnostic "Run clicked"
        $isWeasisOnly = $chkWeasisOnly.Checked

        if (-not $isWeasisOnly -and [string]::IsNullOrWhiteSpace($txtSrc.Text)) {
            [void][System.Windows.Forms.MessageBox]::Show("Select a source folder.", "Validation", "OK", "Warning")
            return
        }
        if ([string]::IsNullOrWhiteSpace($txtDest.Text)) {
            [void][System.Windows.Forms.MessageBox]::Show("Select a destination folder.", "Validation", "OK", "Warning")
            return
        }

        $argsList = New-Object System.Collections.Generic.List[string]

        if (-not $isWeasisOnly) {
            [void]$argsList.Add("-Src")
            [void]$argsList.Add($txtSrc.Text.Trim())
        }

        [void]$argsList.Add("-Dest")
        [void]$argsList.Add($txtDest.Text.Trim())

        [void]$argsList.Add("-Subdir")
        [void]$argsList.Add("IMAGES")

        if ($chkPackage.Checked) {
            [void]$argsList.Add("-PackageWeasis")
        }
        if ($chkVerify.Checked) {
            [void]$argsList.Add("-VerifyDicomdir")
        }
        if ($isWeasisOnly) {
            [void]$argsList.Add("-WeasisOnly")
        }
        if ($chkDeidentify.Checked) {
            [void]$argsList.Add("-Deidentify")
        }

        $cmdPath = $null
        $cmdArgList = @()
        $cmdArgsDisplay = $null

        if (Test-Path -LiteralPath $backendExe -PathType Leaf) {
            $cmdPath = $backendExe
            $cmdArgList = $argsList.ToArray()
        }
        elseif (Test-Path -LiteralPath $backendPs1 -PathType Leaf) {
            $cmdPath = "powershell.exe"
            $head = @("-NoProfile", "-ExecutionPolicy", "Bypass", "-File", $backendPs1)
            $cmdArgList = @($head + $argsList.ToArray())
        }
        else {
            throw "Backend not found. Expected one of: $backendExe or $backendPs1"
        }

        $cmdArgsDisplay = ($cmdArgList | ForEach-Object { ConvertTo-QuotedArgument $_ }) -join " "

        $txtLog.Clear()
        $script:lastDest = $txtDest.Text.Trim()
        $script:completionHandled = $false
        Cleanup-ProcessOutputFiles
        $script:stdoutPath = [System.IO.Path]::GetTempFileName()
        $script:stderrPath = [System.IO.Path]::GetTempFileName()
        $script:progressTotal = 0
        $script:progressProcessed = 0
        $script:progressCopied = 0
        $script:progressDiscovered = 0
        $script:runStart = Get-Date
        $script:discoverySeconds = $null
        $script:scanSeconds = $null
        $script:dicomdirSeconds = $null
        Update-Summary -State "Starting"
        Add-LogLine ("Command: {0} {1}" -f $cmdPath, $cmdArgsDisplay)
        Write-Diagnostic ("ArgCount=" + $cmdArgList.Count + "; Args=" + ($cmdArgList -join " | "))

        $startParams = @{
            FilePath               = $cmdPath
            ArgumentList           = $cmdArgList
            WorkingDirectory       = $scriptDir
            WindowStyle            = 'Hidden'
            PassThru               = $true
            RedirectStandardOutput = $script:stdoutPath
            RedirectStandardError  = $script:stderrPath
        }

        $proc = Start-Process @startParams
        if (-not $proc) {
            throw "Failed to start backend process."
        }

        $script:currentProcess = $proc
        Set-RunningState -Running $true
        Write-Diagnostic "Backend process started"
    }
    catch {
        Write-Diagnostic ("Run handler error: " + $_.Exception.ToString())
        Set-RunningState -Running $false
        [void][System.Windows.Forms.MessageBox]::Show($_.Exception.Message, "Error", "OK", "Error")
    }
})

$btnClose.Add_Click({
    $form.Close()
})

$form.Add_FormClosing({
    if ($script:currentProcess -and -not $script:currentProcess.HasExited) {
        $r = [System.Windows.Forms.MessageBox]::Show(
            "A run is still active. Stop it and close?",
            "Confirm Close",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question
        )

        if ($r -ne [System.Windows.Forms.DialogResult]::Yes) {
            $_.Cancel = $true
            return
        }

        try { $script:currentProcess.Kill() } catch {}
    }
})

$form.Add_FormClosed({
    Cleanup-ProcessOutputFiles
    Write-Diagnostic "GUI closed"
})

[void]$form.ShowDialog()
