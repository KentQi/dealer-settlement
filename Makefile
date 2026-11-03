# dealer-settlement 开发命令
# 全部只依赖 python3 标准库 + requirements.txt，不访问网络。
PYTHON ?= python3

.PHONY: help check demo package clean

help:
	@printf '%s\n' \
	  'make check    校验插件/技能清单、语法编译' \
	  'make demo     三套端到端演示：对账 / L6 复核 / 结构映射' \
	  'make package  构建 dist/dealer-settlement.zip（扁平技能安装包）' \
	  'make clean    清理构建产物与 __pycache__'

check:
	$(PYTHON) tools/validate.py
	$(PYTHON) -m compileall -q skills examples tools

demo:
	$(PYTHON) examples/run_demo.py
	$(PYTHON) examples/review_demo.py
	$(PYTHON) examples/schema_demo.py

package: check
	@mkdir -p dist
	@rm -f dist/dealer-settlement.zip
	cd skills && zip -qr ../dist/dealer-settlement.zip dealer-settlement \
	  -x '*/__pycache__/*' '*.pyc' '*.pyo' '*.DS_Store' '__pycache__/*'
	@echo "打包完成：dist/dealer-settlement.zip"

clean:
	rm -rf dist
	find . -type d -name __pycache__ -prune -exec rm -rf {} + 2>/dev/null || true
