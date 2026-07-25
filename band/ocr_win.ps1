# ocr_win.ps1 — Windows 내장 OCR(Windows.Media.Ocr)로 이미지에서 한국어 텍스트 추출
# ============================================================================
# 외부 서비스·API키·업로드가 전혀 없다. 거래명세서·세금계산서 같은 금융 문서를
# PC 밖으로 내보내지 않기 위해 반드시 이 방식만 쓴다.
#
#   powershell -NoProfile -File ocr_win.ps1 -Path "C:\...\명세서.jpg"
#   → 표준출력으로 인식 텍스트(UTF-8)
param(
  [Parameter(Mandatory=$true)][string]$Path,
  [string]$Lang = 'ko'
)
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

Add-Type -AssemblyName System.Runtime.WindowsRuntime | Out-Null
$asTaskGeneric = ([System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object {
  $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and
  $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1' })[0]

function Await($op, $type) {
  $t = $asTaskGeneric.MakeGenericMethod($type).Invoke($null, @($op))
  $t.Wait(-1) | Out-Null
  $t.Result
}

$null = [Windows.Storage.StorageFile,            Windows.Storage,          ContentType=WindowsRuntime]
$null = [Windows.Graphics.Imaging.BitmapDecoder, Windows.Graphics.Imaging, ContentType=WindowsRuntime]
$null = [Windows.Media.Ocr.OcrEngine,            Windows.Media,            ContentType=WindowsRuntime]
$null = [Windows.Globalization.Language,         Windows.Globalization,    ContentType=WindowsRuntime]

$full = (Resolve-Path -LiteralPath $Path).Path
$file = Await ([Windows.Storage.StorageFile]::GetFileFromPathAsync($full)) ([Windows.Storage.StorageFile])
$strm = Await ($file.OpenAsync([Windows.Storage.FileAccessMode]::Read)) ([Windows.Storage.Streams.IRandomAccessStream])
$dec  = Await ([Windows.Graphics.Imaging.BitmapDecoder]::CreateAsync($strm)) ([Windows.Graphics.Imaging.BitmapDecoder])
$bmp  = Await ($dec.GetSoftwareBitmapAsync()) ([Windows.Graphics.Imaging.SoftwareBitmap])

$engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromLanguage([Windows.Globalization.Language]::new($Lang))
if (-not $engine) { $engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromUserProfileLanguages() }
if (-not $engine) { Write-Error "OCR 엔진 없음 (언어: $Lang)"; exit 2 }

$res = Await ($engine.RecognizeAsync($bmp)) ([Windows.Media.Ocr.OcrResult])
# 줄 단위로 내보낸다 — 표 형태 문서는 줄 구조가 파싱에 중요하다
foreach ($line in $res.Lines) { [Console]::Out.WriteLine($line.Text) }
