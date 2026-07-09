.PHONY: clean-pyc clean-build clean-test help
SHELL := /bin/bash

help:
	@echo "clean - remove all build, test, coverage and Python artifacts"
	@echo "clean-build - remove build artifacts"
	@echo "clean-pyc - remove Python file artifacts"
	@echo "clean-test - remove test and coverage artifacts"
	@echo "lint - check style with flake8"
	@echo "test - run tests with pytest"
	@echo "coverage - check code coverage"
	@echo "dist - package"

clean: clean-build clean-pyc clean-test

clean-build:
	rm -fr build/
	rm -fr dist/
	rm -fr *.egg-info

clean-pyc:
	find . -name '*.pyc' -exec rm -f {} +
	find . -name '*.pyo' -exec rm -f {} +
	find . -name '*~' -exec rm -f {} +
	find . -name '__pycache__' -exec rm -fr {} +

clean-test:
	rm -fr .tox/
	rm -f .coverage
	rm -fr htmlcov/

lint:
	flake8 docx2csv tests --config=./flake8

test:
	pytest --cov=docx2csv --cov-report=term-missing

coverage:
	pytest --cov=docx2csv --cov-report=html
	python -m webbrowser htmlcov/index.html

dist: clean
	python -m build
	ls -l dist
