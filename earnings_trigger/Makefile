VENV := .venv

# Cross-platform venv paths
UNAME := $(shell uname -s 2>/dev/null || echo Windows)
ifeq ($(UNAME),Darwin)
  PYTHON := python3
  PIP := $(VENV)/bin/pip
  PY := $(VENV)/bin/python
else
  PYTHON := python
  PIP := $(VENV)/Scripts/pip
  PY := $(VENV)/Scripts/python
endif

all: init

$(VENV):
	$(PYTHON) -m venv $(VENV)
	$(PIP) install --upgrade pip

install: $(VENV)
	$(PIP) install anthropic google-api-python-client google-auth-httplib2 google-auth-oauthlib yfinance requests

init-db:
	$(PY) db.py

calendar:
	$(PY) calendar_watcher.py

sec:
	$(PY) sec_trigger.py

gmail:
	$(PY) gmail_trigger.py

run-jobs:
	$(PY) job_runner.py

pipeline:
	$(PY) main.py

scan:
	$(PY) main.py --scan-only

status:
	$(PY) main.py --status

purge:
	$(PY) job_runner.py --purge

init: install init-db
	@echo "環境與 DB 已初始化。"

clean:
	rm -rf $(VENV) data/trigger.db

.PHONY: all install init-db calendar sec gmail run-jobs pipeline scan status purge init clean
