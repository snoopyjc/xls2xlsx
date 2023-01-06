#!/usr/bin/env python

"""The setup script."""

from setuptools import setup, find_packages

with open('README.rst') as readme_file:
    readme = readme_file.read()

with open('HISTORY.rst') as history_file:
    history = history_file.read()

requirements = [ 'xlrd', 'openpyxl', 'requests', 'beautifulsoup4', 'Pillow',
        'python-dateutil', 'cssutils', 'webcolors', 'currency-symbols',
        'chardet', 'fonttools', 'PyYAML']

setup_requirements = ['pytest-runner', ]

test_requirements = ['pytest>=3', ]

setup(
    author="Joe Cool",
    author_email='snoopyjc@gmail.com',
    python_requires='>=3.7',
    classifiers=[
        'Development Status :: 4 - Beta',
        'Intended Audience :: Developers',
        'License :: OSI Approved :: MIT License',
        'Natural Language :: English',
        'Programming Language :: Python :: 3',
        'Programming Language :: Python :: 3.7',
        'Programming Language :: Python :: 3.8',
        'Programming Language :: Python :: 3.9',
        'Programming Language :: Python :: 3.10',
    ],
    description="Convert xls file to xlsx",
    entry_points={
        'console_scripts': [
            'xls2xlsx=xls2xlsx.cli:main',
        ],
    },
    install_requires=requirements,
    license="MIT license",
    long_description=readme + '\n\n' + history,
    long_description_content_type='text/x-rst',
    include_package_data=True,
    keywords='xls2xlsx',
    name='xls2xlsx',
    packages=find_packages(include=['xls2xlsx', 'xls2xlsx.*']),
    setup_requires=setup_requirements,
    test_suite='tests',
    tests_require=test_requirements,
    url='https://github.com/snoopyjc/xls2xlsx',
    version='0.2.0',
    zip_safe=False,
)
