#!/bin/sh
# 在已安装 requirements.txt 的 Python 环境中执行；任一步失败即停止。
set -eu
cd "$(dirname "$0")"
python build_inputs.py
python rolling_alpha.py
python merge_panel.py
