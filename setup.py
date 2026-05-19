from pathlib import Path

from setuptools import setup

ROOT = Path(__file__).resolve().parent

setup(
    name="wps-mcp-server",
    version="0.1.0",
    description="A stdio MCP server for controlling WPS Office on Windows via COM.",
    long_description=(ROOT / "README.md").read_text(encoding="utf-8"),
    long_description_content_type="text/markdown",
    py_modules=["wps_server"],
    python_requires=">=3.9",
    install_requires=[
        "pywin32>=306; platform_system == 'Windows'",
    ],
    extras_require={
        "pdf": ["PyMuPDF>=1.23", "pypdf>=4"],
        "dev": ["build>=1.2", "twine>=5"],
    },
    entry_points={
        "console_scripts": [
            "mcp-wps=wps_server:main",
        ],
    },
    license="MIT",
    author="wps-mcp-server contributors",
    url="https://github.com/2365203723/wps-mcp-server",
    project_urls={
        "Homepage": "https://github.com/2365203723/wps-mcp-server",
        "Repository": "https://github.com/2365203723/wps-mcp-server",
        "Issues": "https://github.com/2365203723/wps-mcp-server/issues",
    },
    keywords=["mcp", "wps", "office", "claude", "codex", "windows", "com"],
    classifiers=[
        "Development Status :: 3 - Alpha",
        "Environment :: Win32 (MS Windows)",
        "Intended Audience :: End Users/Desktop",
        "License :: OSI Approved :: MIT License",
        "Operating System :: Microsoft :: Windows",
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.9",
        "Programming Language :: Python :: 3.10",
        "Programming Language :: Python :: 3.11",
        "Programming Language :: Python :: 3.12",
        "Topic :: Office/Business",
    ],
)
