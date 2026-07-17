#!/usr/bin/env bash
# Разовая донастройка ВНУТРИ контейнера: python-зависимости + корпоративный шрифт.
# Запуск: docker compose exec code-server bash deploy/provision.sh
set -e
cd /home/coder/project

echo "==> Python-зависимости"
pip3 install --user -r requirements.txt

echo "==> Шрифт Geologica"
mkdir -p ~/.local/share/fonts
cp "Шрифты (корпоративные)/"*.ttf ~/.local/share/fonts/ 2>/dev/null || true
fc-cache -f

echo "==> Готово. Открой терминал в веб-IDE и запусти:  claude"
echo "    (при первом запуске возможен вход: /login по подписке, либо задан ANTHROPIC_API_KEY)"
