"""Setup configuration for cli-anything-pyaedt package."""

from setuptools import setup, find_namespace_packages

setup(
    name="cli-anything-pyaedt",
    version="1.0.0",
    description="CLI harness for Ansys PyAEDT - Python interface to Ansys Electronics Desktop",
    long_description=open("README.md").read() if __import__("os").path.exists("README.md") else "",
    long_description_content_type="text/markdown",
    author="PyAEDT Team",
    author_email="pyansys.core@ansys.com",
    url="https://github.com/ansys/pyaedt",
    license="MIT",
    packages=find_namespace_packages(include=["cli_anything.*"]),
    install_requires=[
        "click>=8.0.0",
        "prompt-toolkit>=3.0.0",
        "psutil>=5.9.0",
    ],
    extras_require={
        "pyaedt": ["pyaedt>=0.6.0"],
    },
    entry_points={
        "console_scripts": [
            "cli-anything-pyaedt=cli_anything.pyaedt.pyaedt_cli:main",
        ],
    },
    python_requires=">=3.10",
    classifiers=[
        "Development Status :: 4 - Beta",
        "Intended Audience :: Developers",
        "Intended Audience :: Science/Research",
        "License :: OSI Approved :: MIT License",
        "Operating System :: Microsoft :: Windows",
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.10",
        "Programming Language :: Python :: 3.11",
        "Programming Language :: Python :: 3.12",
        "Topic :: Scientific/Engineering",
        "Topic :: Software Development :: Libraries :: Python Modules",
    ],
)
