#!/usr/bin/env python3
"""
Setup script for LlamaDocx package.
"""

import os
from setuptools import setup, find_packages

# Get the long description from the README file
with open(os.path.join(os.path.dirname(__file__), 'README.md'), encoding='utf-8') as f:
    long_description = f.read()

setup(
    name="llamadocx-llamasearch",
    version="0.1.0",
    packages=find_packages(where="src"),
    package_dir={"": "src"},
    description="A powerful toolkit for working with Microsoft Word documents",
    long_description=long_description,
    long_description_content_type="text/markdown",
    author="LlamaSearch AI",
    author_email="nikjois@llamasearch.ai",
    url="https://llamasearch.ai",
    classifiers=[
        "Development Status :: 4 - Beta",
        "Intended Audience :: Developers",
        "License :: OSI Approved :: MIT License",
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.8",
        "Programming Language :: Python :: 3.9",
        "Programming Language :: Python :: 3.10",
        "Programming Language :: Python :: 3.11",
        "Topic :: Office/Business",
        "Topic :: Software Development :: Libraries",
    ],
    python_requires=">=3.8",
    install_requires=[
        "python-docx>=0.8.11",
        "markdown>=3.4.1",
        "beautifulsoup4>=4.11.1",
    ],
    extras_require={
        "dev": [
            "pytest>=7.0.0",
            "pytest-cov>=3.0.0",
            "black>=22.1.0",
            "isort>=5.10.1",
            "flake8>=4.0.1",
            "mypy>=0.931",
        ],
        "convert": [
            "pypandoc>=1.6.0",
            "docx2pdf>=0.1.8",
        ],
        "full": [
            "pypandoc>=1.6.0",
            "docx2pdf>=0.1.8",
            "rapidfuzz>=2.0.0",
            "Pillow>=9.0.0",
            "opencv-python>=4.5.0",
        ],
    },
    entry_points={
        "console_scripts": [
            "llamadocx=llamadocx.cli:cli",
        ],
    },
    include_package_data=True,
    zip_safe=False,
) # Version bump for first release

# Updated in commit 5 - 2025-04-04 17:21:33

# Updated in commit 13 - 2025-04-04 17:21:33

# Updated in commit 21 - 2025-04-04 17:21:34

# Updated in commit 29 - 2025-04-04 17:21:35

# Updated in commit 5 - 2025-04-05 14:30:42

# Updated in commit 13 - 2025-04-05 14:30:42

# Updated in commit 21 - 2025-04-05 14:30:42

# Updated in commit 29 - 2025-04-05 14:30:43

# Updated in commit 5 - 2025-04-05 15:17:09

# Updated in commit 13 - 2025-04-05 15:17:09

# Updated in commit 21 - 2025-04-05 15:17:09

# Updated in commit 29 - 2025-04-05 15:17:10

# Updated in commit 5 - 2025-04-05 15:47:55

# Updated in commit 13 - 2025-04-05 15:47:55

# Updated in commit 21 - 2025-04-05 15:47:55

# Updated in commit 29 - 2025-04-05 15:47:55

# Updated in commit 5 - 2025-04-05 16:53:00

# Updated in commit 13 - 2025-04-05 16:53:00
