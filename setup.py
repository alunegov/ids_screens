from setuptools import setup

PROJECT_NAME = "ids_screens"

setup(
    name=PROJECT_NAME,
    version='0.1',
    description='ids_screens',
    url='https://github.com/alunegov/ids_screens',
    author='Alexander Lunegov',
    author_email='alunegov@gmail.com',
    license='MIT',
    packages=[PROJECT_NAME],
    install_requires=[
        "pywinauto~=0.6.4",
        "PyYAML~=3.12",
    ],
    entry_points={
        "console_scripts": ['{0} = {0}.{0}:main'.format(PROJECT_NAME)],
    }
)
