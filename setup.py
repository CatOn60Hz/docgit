from setuptools import setup

setup(
    name="docgit",
    version="1.2",
    py_modules=["docgit"],
    install_requires=[
        "click",
        "rich",
        "python-docx",
        "pywin32",
    ],
    entry_points={
        "console_scripts": [
            "docgit=docgit:cli",
        ],
    },
)
