from setuptools import setup, find_packages

setup(
    name='scrapingrps',
    version='0.7',
    author='Rafael Perroud Sampaio',
    author_email='rafapsampaio@gmail.com',
    description='Scraping interno',
    long_description=open('README.md').read(),
    long_description_content_type='text/markdown',
    url='https://github.com/rafaelpsampaio/scrapingrps',
    packages=find_packages(),
    install_requires=[
        'selenium',
        'pandas',
        'webdriver_manager',
        'sqlalchemy'
    ],
    classifiers=[
        'Programming Language :: Python :: 3',
        'License :: OSI Approved :: MIT License',
        'Operating System :: OS Independent',
    ],
    python_requires='>=3.6',
)
