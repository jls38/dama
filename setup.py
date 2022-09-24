from setuptools import setup, find_packages

setup(
    name='DAMA',
    version='1.1',
    author='JL SARRAT',
    description='Do things',
    long_description='Do things and other things',
    keywords='do, things',
    python_requires='>=3.10.4',
    packages=find_packages(include=['DAMA', 'DAMA.*']),
    install_requires=[
        'pandas==1.4.2',
        'openpyxl==3.0.9',
        'numpy>=1.22.3',
        'matplotlib>=3.5.2',
        'tk>=8.6.12',
    ],
)