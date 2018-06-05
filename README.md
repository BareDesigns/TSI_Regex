This program was made to extract line by line an EDI transcript to use for
evaluations and AutoHotKey. :fire:

# Pre-Reqs

In order to run this script you will need to have a few things installed:

-   [Python](https://www.python.org)
-   [AutoHotKey](https://www.autohotkey.com)

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

### TODO

There are a few things that still need to be done before full production:

-   Review the code and remove any redundancies.
-   Anything else I think of later...
-   ~~Clean up transfer star line~~
-   ~~Core star line cleanup~~
-   ~~Change '<<CREDIT' term code~~
-   ~~Remove the word **SEMESTER** from the line '<<Fall Semester 2017'~~
-   ~~Fix the spacing after the term year~~
-   **COASTLINE CC** has weird class formatting. Be sure to fix it.
-   Auto-Name the file to the current date
