from setuptools import setup, find_packages

setup(
    name="msiparser-gui",
    version="0.1.0",
    packages=find_packages(),
    install_requires=[
        "PyQt5>=5.15.11",
        "magika>=0.5.1",
    ],
    entry_points={
        "console_scripts": [
            "msiparser-gui=msiparser_gui:main",
        ],
    },
) 