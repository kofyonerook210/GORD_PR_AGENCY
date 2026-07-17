# Деплой на свой VPS — работа через веб-интерфейс

Поднимает **VS Code в браузере (code-server) + Claude Code** на вашем VPS за HTTPS.
Внутри — весь репозиторий GORD как проект: скиллы, `CLAUDE.md`, шаблоны и правила
подхватываются автоматически, Python/LibreOffice/шрифт готовы к сборке отчётов.
Заходите с любого ПК по адресу `https://<домен>` под паролем.

## Требования
- VPS с **Ubuntu/Debian**, публичный IP, ~2 ГБ RAM и ~5 ГБ диска (LibreOffice «тяжёлый»).
- **Домен** с A-записью на IP VPS (нужен для авто-HTTPS). Без домена можно временно по
  IP + HTTP, но это небезопасно.
- Установленные **Docker** и **Docker Compose**.

## Установка Docker (если ещё нет)
```bash
curl -fsSL https://get.docker.com | sh
sudo usermod -aG docker $USER   # перелогиниться после этого
```

## Разворачивание
```bash
# 1. Склонировать репозиторий на VPS
git clone https://github.com/kofyonerook210/GORD_PR_AGENCY.git
cd GORD_PR_AGENCY/deploy

# 2. Настроить переменные
cp .env.example .env
nano .env        # указать DOMAIN, CS_PASSWORD, при желании ANTHROPIC_API_KEY

# 3. Собрать и запустить (первый билд ставит LibreOffice/Node — несколько минут)
docker compose up -d --build

# 4. Донастроить окружение (python-зависимости + шрифт Geologica)
docker compose exec code-server bash deploy/provision.sh
```

## Работа
1. Открыть `https://<домен>` → ввести `CS_PASSWORD` → откроется VS Code в браузере.
2. Меню **Terminal → New Terminal** → выполнить:
   ```bash
   claude
   ```
   При первом запуске — вход: `/login` (по подписке Claude, откроется ссылка для
   авторизации) либо укажите `ANTHROPIC_API_KEY` в `.env` заранее.
3. Ставьте задачи как обычно («собери отчёт по компании X»). Скиллы (`pr-strategy`)
   и правила из `CLAUDE.md`/`memory/` уже в контексте. Готовые файлы — в `Отчеты/`.

## Обновление / GitHub
- Изменения коммитить и пушить из терминала веб-IDE (`git push`). Настройте git-доступ:
  `gh auth login` или personal access token.
- Обновить код деплоя: `git pull && docker compose up -d --build`.

## Безопасность
- `.env` (пароль, ключи) в git не попадает (`deploy/.gitignore`).
- Данные проекта остаются на вашем VPS; наружу уходят только запросы к LLM-провайдеру
  (Claude API/подписка). Пароль входа держите сложным; при желании добавьте в `Caddyfile`
  базовую аутентификацию или ограничение по IP.

## Альтернатива — чат-интерфейс на OpenAI
Если хочется именно «ChatGPT-подобный» веб-чат, а не IDE, вместо code-server можно поднять
**LibreChat** или **Open WebUI** (Docker) с ключом OpenAI/Claude. Но наш пайплайн сборки
(скрипты + шаблоны) под них придётся переносить/оборачивать в инструмент отдельно —
это дополнительная разработка. Скажите — подготовлю такой вариант.
