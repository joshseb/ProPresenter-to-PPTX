#!/bin/bash
# Launcher for ProPresenter 7 → PowerPoint Converter
# Uses the bundled Python 3.13 virtual environment
DIR="$(cd "$(dirname "$0")" && pwd)"
"$DIR/venv/bin/python3.13" "$DIR/pro_to_pptx.py"
