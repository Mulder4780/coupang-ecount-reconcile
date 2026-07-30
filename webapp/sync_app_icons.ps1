param(
    [string]$Source = (Join-Path $PSScriptRoot "brand\csos-app-icon-source.png"),
    [switch]$SkipInstalledApps
)

$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.Drawing

$webRoot = $PSScriptRoot
$root = Split-Path -Parent $webRoot
$docsRoot = Join-Path $root "docs"
$sizes = @(32, 180, 192, 512)
$sourceImage = [System.Drawing.Image]::FromFile($Source)

try {
    foreach ($size in $sizes) {
        $target = Join-Path $webRoot "icon-$size.png"
        $bitmap = New-Object System.Drawing.Bitmap(
            $size,
            $size,
            [System.Drawing.Imaging.PixelFormat]::Format32bppArgb
        )
        try {
            $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
            try {
                $graphics.CompositingMode = [System.Drawing.Drawing2D.CompositingMode]::SourceCopy
                $graphics.CompositingQuality = [System.Drawing.Drawing2D.CompositingQuality]::HighQuality
                $graphics.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
                $graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::HighQuality
                $graphics.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
                $graphics.DrawImage($sourceImage, 0, 0, $size, $size)
            }
            finally {
                $graphics.Dispose()
            }
            $bitmap.Save($target, [System.Drawing.Imaging.ImageFormat]::Png)
        }
        finally {
            $bitmap.Dispose()
        }
    }
}
finally {
    $sourceImage.Dispose()
}

& (Join-Path $webRoot "build_windows_icon.ps1") -Source $Source -Output (Join-Path $webRoot "csos-app.ico") | Out-Null

[System.IO.Directory]::CreateDirectory($docsRoot) | Out-Null
foreach ($size in $sizes) {
    Copy-Item -LiteralPath (Join-Path $webRoot "icon-$size.png") `
        -Destination (Join-Path $docsRoot "icon-$size.png") -Force
}

$sourceHash = (Get-FileHash -LiteralPath $Source -Algorithm SHA256).Hash.ToLowerInvariant()
$revision = "csos-" + $sourceHash.Substring(0, 12)

$manifestPath = Join-Path $docsRoot "manifest.json"
if (Test-Path -LiteralPath $manifestPath) {
    $manifest = Get-Content -LiteralPath $manifestPath -Raw -Encoding UTF8 | ConvertFrom-Json
    foreach ($icon in $manifest.icons) {
        # A previous interrupted update may have left a malformed src.  Never
        # reuse it: the manifest size is the stable source of truth.
        $size = ([string]$icon.sizes).Split("x")[0]
        if ($size -notin @("192", "512")) {
            $size = "192"
        }
        $icon.src = "icon-${size}.png?v=$revision"
    }
    $manifestJson = $manifest | ConvertTo-Json -Depth 8
    [System.IO.File]::WriteAllText(
        $manifestPath,
        $manifestJson + [Environment]::NewLine,
        (New-Object System.Text.UTF8Encoding($false))
    )
}

$serviceWorkerPath = Join-Path $docsRoot "sw.js"
if (Test-Path -LiteralPath $serviceWorkerPath) {
    $serviceWorker = Get-Content -LiteralPath $serviceWorkerPath -Raw -Encoding UTF8
    $serviceWorker = [regex]::Replace(
        $serviceWorker,
        "const CACHE = '[^']+';",
        "const CACHE = 'csos-icon-$($sourceHash.Substring(0, 12))-2026-only';"
    )
    [System.IO.File]::WriteAllText(
        $serviceWorkerPath,
        $serviceWorker.TrimEnd() + [Environment]::NewLine,
        (New-Object System.Text.UTF8Encoding($false))
    )
}

$installedCount = 0
$shortcutCount = 0
if (-not $SkipInstalledApps -and $env:LOCALAPPDATA) {
    $icoSource = Join-Path $webRoot "csos-app.ico"
    $chromeApps = Join-Path $env:LOCALAPPDATA "Google\Chrome\User Data\Default\Web Applications"
    if (Test-Path -LiteralPath $chromeApps) {
        $webAppDirs = Get-ChildItem -LiteralPath $chromeApps -Directory -Filter "_crx_*" -ErrorAction SilentlyContinue
        foreach ($dir in $webAppDirs) {
            $link = Get-ChildItem -LiteralPath $dir.FullName -Filter "Coupang Service Operations System.lnk" `
                -File -ErrorAction SilentlyContinue | Select-Object -First 1
            if (-not $link) {
                continue
            }
            $icoTarget = Join-Path $dir.FullName "Coupang Service Operations System.ico"
            if ((Test-Path -LiteralPath $icoTarget) -and -not (Test-Path -LiteralPath "$icoTarget.legacy")) {
                Copy-Item -LiteralPath $icoTarget -Destination "$icoTarget.legacy" -Force
            }
            Copy-Item -LiteralPath $icoSource -Destination $icoTarget -Force

            $md5 = [System.Security.Cryptography.MD5]::Create()
            try {
                $rawHash = $md5.ComputeHash([System.IO.File]::ReadAllBytes($icoTarget))
                [System.IO.File]::WriteAllBytes("$icoTarget.md5", $rawHash)
            }
            finally {
                $md5.Dispose()
            }
            $installedCount++
        }
    }

    $shortcutRoots = @(
        (Join-Path $env:USERPROFILE "Desktop"),
        (Join-Path $env:APPDATA "Microsoft\Windows\Start Menu\Programs")
    )
    $shell = New-Object -ComObject WScript.Shell
    foreach ($shortcutRoot in $shortcutRoots) {
        if (-not (Test-Path -LiteralPath $shortcutRoot)) {
            continue
        }
        $shortcuts = Get-ChildItem -LiteralPath $shortcutRoot -Recurse -File -Filter "*.lnk" `
            -ErrorAction SilentlyContinue | Where-Object {
                $_.BaseName -eq "Coupang Service Operations System" -or
                $_.BaseName -eq "CSOS 앱 열기"
            }
        foreach ($shortcut in $shortcuts) {
            $item = $shell.CreateShortcut($shortcut.FullName)
            $item.IconLocation = "$icoSource,0"
            $item.Save()
            $shortcutCount++
        }
    }

    $refresh = Join-Path $env:SystemRoot "System32\ie4uinit.exe"
    if (Test-Path -LiteralPath $refresh) {
        Start-Process -FilePath $refresh -ArgumentList "-show" -WindowStyle Hidden -Wait
    }
}

Write-Output "CSOS_ICON_REVISION=$revision INSTALLED=$installedCount SHORTCUTS=$shortcutCount"
