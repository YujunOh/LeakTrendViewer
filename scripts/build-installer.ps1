param(
    [string]$Configuration = "Release",
    [string]$Runtime = "win-x64"
)

$ErrorActionPreference = "Stop"

$ProjectRoot = Split-Path -Parent $PSScriptRoot
$PublishDir = Join-Path $ProjectRoot "publish\$Runtime"
$DistDir = Join-Path $ProjectRoot "dist"
$InstallerScript = Join-Path $ProjectRoot "installer\LeakTrendViewer.iss"

Write-Host "[1/3] Publish app..." -ForegroundColor Cyan
dotnet publish (Join-Path $ProjectRoot "LeakTrendViewer.csproj") -c $Configuration -r $Runtime --self-contained true /p:PublishSingleFile=true /p:PublishTrimmed=false -o $PublishDir

if (-not (Test-Path $PublishDir)) {
    throw "Publish output not found: $PublishDir"
}

Write-Host "[2/3] Check Inno Setup..." -ForegroundColor Cyan
$isccCandidates = @(
    "C:\Program Files (x86)\Inno Setup 6\ISCC.exe",
    "C:\Program Files\Inno Setup 6\ISCC.exe"
)

$iscc = $isccCandidates | Where-Object { Test-Path $_ } | Select-Object -First 1

if (-not $iscc) {
    Write-Host "ISCC.exe not found. Inno Setup 설치 후 아래 명령 실행:" -ForegroundColor Yellow
    Write-Host "  `"C:\Program Files (x86)\Inno Setup 6\ISCC.exe`" `"$InstallerScript`""
    Write-Host "Publish 결과물만 생성 완료: $PublishDir" -ForegroundColor Green
    exit 0
}

Write-Host "[3/3] Build installer..." -ForegroundColor Cyan
& $iscc $InstallerScript | Out-Host

if (-not (Test-Path $DistDir)) {
    throw "Installer output folder not found: $DistDir"
}

Write-Host "완료. 설치파일 폴더: $DistDir" -ForegroundColor Green
