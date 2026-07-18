#!/bin/bash
set -e

# Отримуємо абсолютний шлях до папки, де зараз лежить цей скрипт, phones.py та finder.py
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

echo "📦 1. Створюємо venv поруч із файлами..."
python3 -m venv "$SCRIPT_DIR/venv"
"$SCRIPT_DIR/venv/bin/pip" install --quiet python-docx

echo "⚙️ 2. Створюємо запускні файли в ~/.local/bin..."
mkdir -p "$HOME/.local/bin"

# Створюємо команду 'phones'
cat << RUNEOF > "$HOME/.local/bin/phones"
#!/bin/bash
"$SCRIPT_DIR/venv/bin/python" "$SCRIPT_DIR/phones.py" "\$@"
RUNEOF
chmod +x "$HOME/.local/bin/phones"

# Створюємо команду 'finder'
cat << RUNEOF > "$HOME/.local/bin/finder"
#!/bin/bash
"$SCRIPT_DIR/venv/bin/python" "$SCRIPT_DIR/finder.py" "\$@"
RUNEOF
chmod +x "$HOME/.local/bin/finder"

echo "✅ Готово! Перезапусти термінал або виконай: source ~/.profile (або source ~/.bashrc)"
echo "Тепер команди 'phones' та 'finder' доступні всюди!"