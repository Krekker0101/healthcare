param()

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$python = Get-Command python -ErrorAction SilentlyContinue
if (-not $python) { $python = Get-Command py -ErrorAction SilentlyContinue }
if (-not $python) { throw "Python не найден. Запустите build.py вручную в среде с Python 3." }

& $python.Source (Join-Path $root "build.py")
if ($LASTEXITCODE -ne 0) { throw "Сборка завершилась с ошибкой: $LASTEXITCODE" }
