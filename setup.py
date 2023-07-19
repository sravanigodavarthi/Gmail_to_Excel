import setuptools

setuptools.setup(
    name="Gmail_to_excel_Package",
    version="1.0.0",
    description="Script to retrieve specific email messages from Gmail and write them to an Excel file",
    packages=setuptools.find_packages(),
    install_requires=[
        "pandas",
        "IMAPClient"
    ],
    )
