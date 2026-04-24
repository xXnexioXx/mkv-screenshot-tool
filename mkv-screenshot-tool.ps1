# ============================================================
#  MKV Screenshot-Tool
#  Requirements: ffmpeg.exe + ffprobe.exe in the same folder
#  Enter your ImgBB API key below before running!
#  made with <3 by nexio
# ============================================================

# ── Auto-restart in PowerShell 7 if opened in PS5 ───────────
if ($PSVersionTable.PSVersion.Major -lt 7) {
    $ps7Path = "$env:ProgramFiles\PowerShell\7\pwsh.exe"
    if (Test-Path $ps7Path) {
        Write-Host "PS5 detected - restarting in PowerShell 7..." -ForegroundColor Yellow
        Start-Process $ps7Path -ArgumentList "-ExecutionPolicy RemoteSigned -File `"$PSCommandPath`"" -Wait
    } else {
        Write-Host "ERROR: PowerShell 7 not found at:" -ForegroundColor Red
        Write-Host "  $ps7Path" -ForegroundColor Yellow
        Write-Host "Please install PS7 from: https://aka.ms/powershell" -ForegroundColor Yellow
        Read-Host "Press Enter to exit"
    }
    exit
}

try {

# ── Configuration ────────────────────────────────────────────
$API_KEY    = "YOUR_IMGBB_API_KEY"   # <-- enter your ImgBB API key from here https://api.imgbb.com/
$SCRIPT_DIR = $PSScriptRoot
$FFMPEG     = Join-Path $SCRIPT_DIR "ffmpeg.exe"
$FFPROBE    = Join-Path $SCRIPT_DIR "ffprobe.exe"

# ── Header ───────────────────────────────────────────────────
Write-Host "================================" -ForegroundColor Cyan
Write-Host "   MKV Screenshot Tool"          -ForegroundColor Cyan
Write-Host "================================" -ForegroundColor Cyan
Write-Host "  Folder : $SCRIPT_DIR"          -ForegroundColor DarkGray
Write-Host "  PS Ver : $($PSVersionTable.PSVersion)" -ForegroundColor DarkGray

# ── Step 0: Remove old .txt files ────────────────────────────
$oldTxtFiles = Get-ChildItem -Path $SCRIPT_DIR -Filter "*.txt"
if ($oldTxtFiles) {
    Write-Host ""
    Write-Host "Removing old .txt files..." -ForegroundColor DarkYellow
    foreach ($file in $oldTxtFiles) {
        Remove-Item $file.FullName -Force
        Write-Host "  Deleted: $($file.Name)" -ForegroundColor DarkGray
    }
}

# ── Step 1: Check for ffmpeg and ffprobe ─────────────────────
Write-Host ""
$missingTools = @()
if (-not (Test-Path $FFMPEG))  { $missingTools += "ffmpeg.exe" }
if (-not (Test-Path $FFPROBE)) { $missingTools += "ffprobe.exe" }

if ($missingTools.Count -gt 0) {
    Write-Host "ERROR: The following tools are missing from the script folder:" -ForegroundColor Red
    $missingTools | ForEach-Object { Write-Host "  - $_" -ForegroundColor Red }
    Write-Host ""
    Write-Host "Please place ffmpeg.exe and ffprobe.exe in:" -ForegroundColor Yellow
    Write-Host "  $SCRIPT_DIR" -ForegroundColor Yellow
    Write-Host ""
    Read-Host "Press Enter to exit"
    exit 1
}

Write-Host "ffmpeg + ffprobe found." -ForegroundColor Green

# ── Step 2: Find MKV file ────────────────────────────────────
$mkvFile = Get-ChildItem -Path $SCRIPT_DIR -Filter "*.mkv" | Select-Object -First 1

if (-not $mkvFile) {
    Write-Host ""
    Write-Host "ERROR: No .mkv file found in folder:" -ForegroundColor Red
    Write-Host "  $SCRIPT_DIR" -ForegroundColor Yellow
    Write-Host ""
    Read-Host "Press Enter to exit"
    exit 1
}

$inputFile = $mkvFile.FullName
$baseName  = $mkvFile.BaseName   # Original filename preserved — all characters supported

Write-Host "File found: $($mkvFile.Name)" -ForegroundColor Green

# ── Step 3: Analyze video streams ────────────────────────────
# Query all video streams and select the one with the highest bitrate.
# Falls back to container duration if stream-level data is unavailable.
Write-Host "Analyzing video streams..." -ForegroundColor DarkCyan

$rawStreams = (& $FFPROBE -v error `
    -select_streams v `
    -show_entries stream=index,bit_rate,duration `
    -of csv=p=0 `
    "$inputFile") | Where-Object { $_ -match '\d' }

if (-not $rawStreams) {
    Write-Host ""
    Write-Host "ERROR: No video streams found or file is unreadable!" -ForegroundColor Red
    Read-Host "Press Enter to exit"
    exit 1
}

$bestStreamIndex = 0
$bestBitrate     = -1
$bestDuration    = 0

foreach ($line in $rawStreams) {
    $parts = $line -split ','
    if ($parts.Count -lt 3) { continue }

    $streamIndex = $parts[0].Trim()
    $bitrate     = $parts[1].Trim()
    $duration    = $parts[2].Trim()

    if ($bitrate  -eq 'N/A' -or $bitrate  -eq '') { $bitrate  = 0 }
    if ($duration -eq 'N/A' -or $duration -eq '') { $duration = 0 }

    $bitrateInt  = [long]$bitrate
    $durationInt = [int]($duration -split '\.')[0]

    Write-Host "  Stream $streamIndex : Bitrate=$bitrateInt bps | Duration=$durationInt sec" -ForegroundColor DarkGray

    if ($bitrateInt -gt $bestBitrate) {
        $bestBitrate     = $bitrateInt
        $bestStreamIndex = [int]$streamIndex
        $bestDuration    = $durationInt
    }
}

# Fallback: read duration from container header if stream-level duration was 0
if ($bestDuration -lt 10) {
    Write-Host "  Stream duration unavailable, reading from container..." -ForegroundColor DarkYellow
    $rawDurationLines = (& $FFPROBE -v error -show_entries format=duration -of default=nw=1:nk=1 "$inputFile") | Where-Object { $_ -match '\d' }
    $rawDuration = $rawDurationLines | Select-Object -First 1
    if ($rawDuration) {
        $bestDuration = [int]($rawDuration -split '\.')[0]
        Write-Host "  Container duration: $bestDuration sec" -ForegroundColor DarkYellow
    }
    $bestStreamIndex = 0
}

$totalDuration = $bestDuration
$streamIndex   = $bestStreamIndex

Write-Host "Using stream : $streamIndex" -ForegroundColor Green

if ($totalDuration -lt 10) {
    Write-Host ""
    Write-Host "ERROR: Video is shorter than 10 seconds ($totalDuration sec) - aborting." -ForegroundColor Red
    Read-Host "Press Enter to exit"
    exit 1
}

$totalMinutes = [math]::Floor($totalDuration / 60)
$totalSeconds = $totalDuration % 60
Write-Host "Duration     : $totalDuration seconds ($totalMinutes min $totalSeconds sec)" -ForegroundColor Green

# ── Step 4: Calculate screenshot timestamps ──────────────────
# Screenshots are distributed evenly between 10% and 90% of the video.
# Dividing the range into 6 intervals places 5 evenly spaced points between them.
$rangeStart    = [int]($totalDuration * 0.10)
$rangeEnd      = [int]($totalDuration * 0.90)
$range         = $rangeEnd - $rangeStart
$interval      = [int]($range / 6)

Write-Host "Range        : $rangeStart - $rangeEnd sec | Interval: $interval sec" -ForegroundColor DarkCyan

# ── Step 5: Capture screenshots ──────────────────────────────
# For each position, up to 3 frames are captured (1 second apart).
# The frame with the largest file size is kept — black/blank frames are typically smallest.

function Invoke-FFmpeg {
    param([string[]]$Arguments)
    $errFile = "$env:TEMP\ffmpeg_err.txt"
    & $FFMPEG @Arguments 2>$errFile
    return $LASTEXITCODE
}

$capturedFrames = @{}

for ($i = 1; $i -le 5; $i++) {
    $targetTimestamp = $rangeStart + $interval * $i

    Write-Host ""
    Write-Host "==========================" -ForegroundColor Yellow
    Write-Host " Screenshot $i/5 at $targetTimestamp sec" -ForegroundColor Yellow
    Write-Host "==========================" -ForegroundColor Yellow

    $bestFrameSize = 0
    $bestFramePath = $null
    $bestFrameTimestamp = $null

    for ($attempt = 0; $attempt -lt 3; $attempt++) {
        $currentTimestamp = $targetTimestamp + $attempt

        $hh = ([int][math]::Floor($currentTimestamp / 3600)).ToString("D2")
        $mm = ([int][math]::Floor(($currentTimestamp % 3600) / 60)).ToString("D2")
        $ss = ([int]($currentTimestamp % 60)).ToString("D2")

        $outputFile = Join-Path $SCRIPT_DIR "$baseName-snapshot-$hh-$mm-$ss.png"

        Write-Host "  Attempt $($attempt+1)/3 at $hh`:$mm`:$ss ..." -ForegroundColor DarkGray

        $ffmpegArgs = @(
            "-ss", "$currentTimestamp",
            "-i", "$inputFile",
            "-map", "0:$streamIndex",
            "-frames:v", "1",
            "-q:v", "1",
            "$outputFile",
            "-y"
        )

        $exitCode = Invoke-FFmpeg -Arguments $ffmpegArgs

        if ($exitCode -ne 0 -or -not (Test-Path $outputFile)) {
            Write-Host "  Failed to capture frame. (Exit: $exitCode)" -ForegroundColor Red
            continue
        }

        $fileSize = (Get-Item $outputFile).Length
        Write-Host "  Frame size: $([math]::Round($fileSize / 1KB, 1)) KB" -ForegroundColor DarkGray

        if ($fileSize -gt $bestFrameSize) {
            if ($bestFramePath -and (Test-Path $bestFramePath) -and $bestFramePath -ne $outputFile) {
                Remove-Item $bestFramePath -Force
            }
            $bestFrameSize      = $fileSize
            $bestFramePath      = $outputFile
            $bestFrameTimestamp = "$hh-$mm-$ss"
        } else {
            Remove-Item $outputFile -Force
        }
    }

    if ($bestFramePath) {
        Write-Host "  Best frame : $(Split-Path -Leaf $bestFramePath) ($([math]::Round($bestFrameSize / 1KB, 1)) KB)" -ForegroundColor Green
        $capturedFrames[$i] = @{ Path = $bestFramePath; Timestamp = $bestFrameTimestamp }
    } else {
        Write-Host "  ERROR: Could not capture a valid frame for position $i!" -ForegroundColor Red
    }
}

# ── Step 6: Upload to ImgBB ──────────────────────────────────
Write-Host ""
Write-Host "================================" -ForegroundColor Cyan
Write-Host " Uploading to ImgBB..."          -ForegroundColor Cyan
Write-Host "================================" -ForegroundColor Cyan

$uploadedUrls = @{}
$failedUploads = @()

for ($i = 1; $i -le 5; $i++) {
    if (-not $capturedFrames.ContainsKey($i)) {
        Write-Host "Frame $i missing, skipping upload..." -ForegroundColor DarkYellow
        $failedUploads += $i
        continue
    }

    $imagePath  = $capturedFrames[$i].Path
    $timestamp  = $capturedFrames[$i].Timestamp
    $imgbbName  = "$baseName [$timestamp]"

    Write-Host "Uploading ($i/5): $(Split-Path -Leaf $imagePath)..." -ForegroundColor White
    Write-Host "  ImgBB name: $imgbbName" -ForegroundColor DarkGray

    try {
        $response = Invoke-RestMethod `
            -Uri "https://api.imgbb.com/1/upload?key=$API_KEY" `
            -Method Post `
            -Form @{ image = Get-Item $imagePath; name = $imgbbName }

        if ($response.success -eq $true) {
            $url = $response.data.url
            $uploadedUrls[$i] = $url
            Write-Host "  OK: $url" -ForegroundColor Green
            Remove-Item $imagePath -Force
            Write-Host "  PNG deleted." -ForegroundColor DarkGray
        } else {
            Write-Host "  ERROR: ImgBB reported failure." -ForegroundColor Red
            $failedUploads += $i
        }
    } catch {
        Write-Host "  ERROR: $($_.Exception.Message)" -ForegroundColor Red
        $failedUploads += $i
    }
}

# ── Step 7: Validate results ─────────────────────────────────
Write-Host ""
if ($failedUploads.Count -gt 0) {
    Write-Host "WARNING: The following screenshots failed to upload:" -ForegroundColor Red
    $failedUploads | ForEach-Object { Write-Host "  - Screenshot $_" -ForegroundColor Red }
    Write-Host ""
    Write-Host "Output .txt will NOT be created as it would be incomplete." -ForegroundColor Red
    Write-Host ""
    Read-Host "Press Enter to exit"
    exit 1
}

# ── Step 8: Write BBCode output file ─────────────────────────
Write-Host "================================" -ForegroundColor Cyan
Write-Host " Writing output file..."         -ForegroundColor Cyan
Write-Host "================================" -ForegroundColor Cyan

$outputFile = Join-Path $SCRIPT_DIR "$baseName.txt"

$u1 = $uploadedUrls[1]
$u2 = $uploadedUrls[2]
$u3 = $uploadedUrls[3]
$u4 = $uploadedUrls[4]
$u5 = $uploadedUrls[5]

$bbCode = "[center]`n[url=$u1][img=500]$u1[/img][/url] [url=$u2][img=500]$u2[/img][/url]`n`n[url=$u3][img=500]$u3[/img][/url] [url=$u4][img=500]$u4[/img][/url]`n`n[url=$u5][img=500]$u5[/img][/url]`n[/center]`n`n`n`n`n[right][url=https://github.com/xXnexioXx/mkv-screenshot-tool][size=4]Created by mkv-screenshot-tool[/size][/url][/right]"

# Write UTF-8 without BOM to avoid stray characters in forum editors
$utf8NoBOM = New-Object System.Text.UTF8Encoding $false
[System.IO.File]::WriteAllText($outputFile, $bbCode, $utf8NoBOM)

Write-Host ""
Write-Host "================================" -ForegroundColor Green
Write-Host " Done!"                           -ForegroundColor Green
Write-Host " Output: $baseName.txt"           -ForegroundColor Green
Write-Host "================================" -ForegroundColor Green
Write-Host ""
Read-Host "Press Enter to exit"

} catch {
    Write-Host ""
    Write-Host "======================================" -ForegroundColor Red
    Write-Host " UNEXPECTED ERROR:"                    -ForegroundColor Red
    Write-Host " $($_.Exception.Message)"              -ForegroundColor Red
    Write-Host " Line: $($_.InvocationInfo.ScriptLineNumber)" -ForegroundColor Red
    Write-Host "======================================" -ForegroundColor Red
    Write-Host ""
    Read-Host "Press Enter to exit"
    exit 1
}
