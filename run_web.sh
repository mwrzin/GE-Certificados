#!/usr/bin/env bash
set -euo pipefail

cd "$(dirname "$0")"
python3 -m venv .venv_certificados
source .venv_certificados/bin/activate
python -m pip install --upgrade pip
python -m pip install -r requirements_certificados.txt
python certificados_web.py
