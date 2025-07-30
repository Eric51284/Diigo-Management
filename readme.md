# This repo contains code for organizing downloads from diigo.com
## Inputs
  - .docx versions of diigo outliner exports
    - This is necessary to retain any outline structure that has been created in diigo outliners
    - However, most of the links will be rendered useless for sharing, because they've been converted to diigo outliner relative links
  - .csv output of entire diigo library
    - This output preserves the original link addresses, so is used to correct the problematic links from outliner export
## Outputs
  - process_diigo_doc.py will use the hierarchically structured .docx file (from the outliner export)
    - extract the header information from the outliner
    - extract dates if they've been included in the outliner
    - (Note: altered format - missing dates, etc. will result in incorrect output)
    - Correct links by matching with .csv file
    - Return output in .xlsx file
## Running the .py file
  - Use diigo_processing.ipynb to specify file paths and run the program