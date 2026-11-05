# dealer-settlement 开发命令
# 全部只依赖 python3 标准库 + requirements.txt，不访问网络。
PYTHON ?= python3

.PHONY: help check demo package clean

help:
	@printf '%s\n' \
	  'make check    校验插件/技能清单、语法编译' \
	  'make demo     四套端到端演示：对账 / L6 复核 / 结构映射 / 确认单定位' \
	  'make package  构建 dist/ 下两个包：技能安装包 + 完整源码包（UTF-8 文件名）' \
	  'make clean    清理构建产物与 __pycache__'

check:
	$(PYTHON) tools/validate.py
	$(PYTHON) -m compileall -q skills examples tools

demo:
	$(PYTHON) examples/run_demo.py
	$(PYTHON) examples/review_demo.py
	$(PYTHON) examples/schema_demo.py
	$(PYTHON) examples/confirm_demo.py

package: check
	$(PYTHON) tools/package_zip.py

clean:
	rm -rf dist
	find . -type d -name __pycache__ -prune -exec rm -rf {} + 2>/dev/null || true
