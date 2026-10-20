.PHONY: check index test

check:
	python -m qa_cases validate

index:
	python -m qa_cases index

test:
	python -m unittest discover -s tests -v
