from setuptools import setup

with open("README.md", "r") as fh:
    long_description = fh.read()

setup(
    name="Markdown2docx",
    version="0.3.0",
    description="Convert Markdown to docx with token substitution and command output substitution",
    long_description=long_description,
    long_description_content_type="text/markdown",
    author="Jeremy Lee",
    author_email="jlee2.71818@gmail.com",
    license="MIT",
    py_modules=["Markdown2docx", "PreprocessMarkdown2docx"],
    package_dir={"": "src"},
    install_requires=[
        "bs4>=0.0.1",
        "pillow",
        "python-docx>=0.8.11",
        "markdown2>=2.4.2",
    ],
    extras_require={
        "dev": [
            "pytest>=3.7",
        ],
    },
    keywords="convert, markdown, docx, merge, tokens, substitution, development",
    classifiers=[
        "Development Status :: 1 - Planning",
        "Intended Audience :: Science/Research",
        "License :: OSI Approved :: MIT License",
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.6",
        "Programming Language :: Python :: 3.7",
        "Programming Language :: Python :: 3.8",
        "Programming Language :: Python :: 3.9",
        "Programming Language :: Python :: 3.10",
        "Operating System :: POSIX :: Linux",
        "Topic :: Text Processing :: Markup :: Markdown",
        "Topic :: Text Processing :: Markup :: reStructuredText",
        "Topic :: Documentation",
    ],
)
