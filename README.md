This program was made to extract the necessary data for TSI purposes from an EDI
trsnscript

# Pre-Reqs

In order to run this script you will need to have a few things installed:

-   [Python](https://www.python.org)

Be sure to add Python to %path% in order to run it straight from the command
line.

## Python Modules

Once Python is installed, you will need to run the following from the command
prompt:

    pip install python-docx

This will install the docx module needed to print fancy things like **bold** or
*italics*.

You can read more about Python-docx
[here](https://python-docx.readthedocs.io/en/latest).

### Potential Issues

Depending on your windows build, you may have to make a word document titled
"default.docx" and put the path in the codebase on line 6 of the code.

This will allow the code to run properly.
