#!/bin/bash

# Автоматический установщик манифеста надстройки для Word на Mac

# Получаем имя пользователя
USERNAME=$(whoami)

# Задаем путь к папке WEF для Word
WEF_PATH="/Users/$USERNAME/Library/Containers/com.microsoft.Word/Data/Documents/wef"

# URL манифеста
MANIFEST_URL="https://raw.githubusercontent.com/daswer123/openrouter-word.github.io/refs/heads/main/manifest_mac.xml"

# Имя файла манифеста
MANIFEST_FILE="manifest_mac.xml"

# Проверяем, существует ли папка WEF, если нет - создаем
if [ ! -d "$WEF_PATH" ]; then
    echo "Создаем папку WEF..."
    mkdir -p "$WEF_PATH"
fi

# Скачиваем манифест
echo "Скачиваем манифест..."
curl -L "$MANIFEST_URL" -o "$WEF_PATH/$MANIFEST_FILE"

# Проверяем, успешно ли скачался файл
if [ $? -eq 0 ]; then
    echo "Манифест успешно установлен в: $WEF_PATH/$MANIFEST_FILE"
    echo "Пожалуйста, перезапустите Word и проверьте наличие надстройки в меню Главная > Надстройки"
else
    echo "Ошибка при скачивании манифеста"
    exit 1
fi

# Проверяем, запущен ли Word и предлагаем перезапустить
if pgrep -f "Microsoft Word" > /dev/null; then
    echo "Word запущен. Закройте и откройте Word заново для применения изменений."
fi

exit 0
