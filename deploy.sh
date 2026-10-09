#!/usr/bin/env sh
set -eu

cd "$(dirname "$0")"

if ! command -v docker >/dev/null 2>&1; then
  echo "Docker CLI не найден. Установите и запустите Docker Desktop." >&2
  exit 1
fi

PORT="${PORT:-8088}"
case "$PORT" in
  ""|*[!0-9]*)
    echo "PORT должен быть числом от 1 до 65535." >&2
    exit 1
    ;;
esac
if [ "$PORT" -lt 1 ] || [ "$PORT" -gt 65535 ]; then
  echo "PORT должен быть числом от 1 до 65535." >&2
  exit 1
fi
export PORT

docker compose up --build --detach
printf 'Приложение доступно на всех интерфейсах, порт %s.\n' "$PORT"
HOST_IP="$(ip -4 route get 1.1.1.1 2>/dev/null | awk '{for (i = 1; i <= NF; i++) if ($i == "src") {print $(i + 1); exit}}' || true)"
if [ -n "$HOST_IP" ]; then
  printf 'Адрес для подключения по сети: http://%s:%s\n' "$HOST_IP" "$PORT"
else
  printf 'Для подключения по сети используйте http://<IP-компьютера>:%s (узнать IP: ipconfig или ip addr).\n' "$PORT"
fi
