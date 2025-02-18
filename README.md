# OCLC-Weeding-Script

This script is designed to convert raw OCLC WMS holding output into human-usable, annotated, formatted sheets of weeding candidates ready for review.

Candidates records are
- Filtered by year of last circulation (this is configurable with a dropdown in the GUI)
- Organized into section groups
- Automatically checked against the other holdings in-state using the Worldcat API (the state is hardcoded in as Florida at line 153; change this in the source code for your own library if necessary)
- Written out to user-friendly .xlsx files

IN ORDER TO RUN THIS SCRIPT YOU MUST HAVE THE FOLLOWING IN THE SAME DIRECTORY
- A file called "credentials.dat" containing your wskey client ID and client secret from OCLC
- A directory called "input" containing at least one OCLC ".xls" record file (actually a tab-delimited text file that they give the .xls suffix for some reason).
