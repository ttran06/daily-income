#!/usr/bin/env python
# -*- encoding: utf-8 -*-

from setuptools import setup


setup(
    name="daily-income",
    version="0.10",
    description="Generate Excel sheet to record business daily income for Super Nails",
    packages=["daily_income"],
    install_requires=[
        "pandas>=1.4.2",
        "xlsxwriter>=3.0.3",
    ],  # external packages acting as dependencies
    entry_points={"console_scripts": ["daily_income=daily_income.__main__:main"]},
)
