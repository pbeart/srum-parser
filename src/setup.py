import setuptools

setuptools.setup(
    name="srum_parser",
    version="0.1.0",
    author="Patrick Beart, Kristian Nymann Jakobsen",
    author_email="unknown, kristian@nymann.dev",
    description="A tool to assist with parsing the SRUDB.dat ESE database",
    url="https://github.com/pbeart/srum-parser",
    packages=setuptools.find_packages(),
    classifiers=[
        "Programming Language :: Python :: 3",
        "Operating System :: OS Independent",
    ],
    entry_points={
        'console_scripts': ['parse-srum = srum_parser.console_scripts:main',]
    },
    include_package_data=False,
    install_requires=['cClick, CppHeaderParser, libesedb-python, pytest'])
