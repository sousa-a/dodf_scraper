from setuptools import setup, find_packages

setup(
    name="dodf_scraper",
    version="0.1.0",
    packages=find_packages(),
    install_requires=[
        "pandas",
        "selenium",
        "openpyxl",
    ],
    entry_points={
        "console_scripts": [
            "dodf_scraper=main:main",
        ],
    },
    author="Alessandro Oliveira de Sousa",
    description="Scraper para extrair Extratos de Notas de Empenho do Diário Oficial do Distrito Federal (DODF)",
)
