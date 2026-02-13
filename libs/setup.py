from setuptools import setup, find_packages

setup(
    name="iasuite-common",
    version="1.0.0",
    description="Shared utilities for Invoice Automation Suite",
    packages=find_packages(),
    python_requires=">=3.10",
    install_requires=[
        "pyyaml>=6.0",
    ],
)
