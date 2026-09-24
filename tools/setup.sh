#!/usr/bin/env bash
# Установка зависимостей для скриптов в tools/ (Python 3.10+).
# Запуск: bash tools/setup.sh
set -euo pipefail

PKGS="python-docx openpyxl pypdf pdfplumber pypdfium2 xlrd"

if pip install --quiet --user $PKGS 2>/dev/null; then
  :
else
  # В системах с «externally managed» Python (Debian 12+) нужен этот флаг
  pip install --quiet --user --break-system-packages $PKGS
fi

python3 -c "import docx, openpyxl, pypdf, pdfplumber, pypdfium2, xlrd; print('Зависимости установлены')"
