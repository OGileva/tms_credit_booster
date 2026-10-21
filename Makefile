.PHONY: check test

PYTHON ?= $(if $(wildcard .venv/bin/python),.venv/bin/python,python3)

check:
	$(PYTHON) -m qa_cases validate

test:
	$(PYTHON) -m unittest discover -s tests -v
