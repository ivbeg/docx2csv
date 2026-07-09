"""Backward-compatible setup.py shim.

All declarative configuration lives in pyproject.toml. This file exists
for tools that still invoke `python setup.py` directly.
"""
from setuptools import setup

setup()
