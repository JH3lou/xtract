from setuptools import setup, find_packages

setup(
    name="TaxOptConverter",
    version="1.0.0",
    packages=find_packages(),
    install_requires=[
        "customtkinter",
        "pandas",
        "openpyxl",
        "tk",
    ],
    entry_points={
        "gui_scripts": [
            "taxopt = app.app:main",
        ]
    },
)
