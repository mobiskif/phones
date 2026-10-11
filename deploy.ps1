param(
    [ValidateRange(1, 65535)]
    [int]$Port = 8090
)

$ErrorActionPreference = "Stop"
$projectDirectory = $PSScriptRoot

$docker = Get-Command docker -ErrorAction SilentlyContinue
if (-not $docker) {
    throw "Docker CLI не найден. Установите Docker Desktop и включите его."
}

foreach ($fileName in @("pass.txt", "structure.xlsx")) {
    $filePath = Join-Path $projectDirectory $fileName
    if (-not (Test-Path -LiteralPath $filePath -PathType Leaf)) {
        throw "Не найден файл '$fileName' в папке проекта: $projectDirectory"
    }
}

Push-Location $projectDirectory
try {
    & $docker.Source info *> $null
    if ($LASTEXITCODE -ne 0) {
        throw "Docker Desktop не запущен или Docker Engine недоступен. Запустите Docker Desktop и повторите попытку."
    }

    & $docker.Source compose version *> $null
    if ($LASTEXITCODE -ne 0) {
        throw "Команда 'docker compose' недоступна. Обновите Docker Desktop до версии с поддержкой Compose."
    }

    $env:PORT = "$Port"
    & $docker.Source compose up --build --detach
    if ($LASTEXITCODE -ne 0) {
        throw "Не удалось развернуть приложение через Docker Compose (код ошибки: $LASTEXITCODE)."
    }
}
finally {
    Pop-Location
}

Write-Host ""
Write-Host "Приложение доступно на этом компьютере: http://localhost:$Port"

$lanAddresses = @(
    Get-NetIPConfiguration |
        Where-Object { $_.NetAdapter.Status -eq "Up" -and $_.IPv4DefaultGateway } |
        ForEach-Object { $_.IPv4Address.IPAddress } |
        Where-Object { $_ -and $_ -notlike "169.254.*" }
)

if ($lanAddresses.Count -gt 0) {
    foreach ($address in ($lanAddresses | Select-Object -Unique)) {
        Write-Host "Адрес для устройств в локальной сети: http://${address}:$Port"
    }
    Write-Host "Если страница недоступна по сети, разрешите входящие подключения на порт $Port в брандмауэре Windows."
}
else {
    Write-Host "Не удалось определить IP-адрес локальной сети. Выполните 'ipconfig', чтобы узнать адрес компьютера."
}
