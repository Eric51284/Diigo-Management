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

## For New Articles Only
 - run just the cells 9 and 10 in diigo_processing.ipynb
   - to set input and output file paths
     - input path is .docx file containing export from diigo outliner with only new files included
     - output path is .xlsx file that will contain fetched dates, along with article titles and links
   - and then will run NewArticles.py that will extract the dates and create the .xlsx file
# UPDATE 2026-02-14
## Transitioned from diigo to raindrop.io
 - imported all diigo bookmarks into raindrop
 - use expand_redirects.py to correct 'flip.it' shortcuts to full urls
 - raindroptagger.py created to collect pub dates and word counts from raindrop.io exported csv files
 - to run, use `rdtagger` snippet in an ipynb cell
 - Once updated csv (with pub dates and wordcounts) is obtained, ~can run `add_raindrop_to_outline.py`~ (this was intended to automate categorization for outlines, but isn't working as well as the original - which was generated using poe.com)
 - MORE FUNCTIONAL APPROACH
   - in raindrop.io, add tags like `_outl:___` with the roman numeral and letter for the appropriate outline section
   - export relevant files to .csv
   - run `add_outl_articles.py` on the exported csv to include in the current html
     - To use it again in the future, just replace the contents of outl.csv with the new export and run: 
  
  `python "Python/add_outl_articles.py"`
   - copy and post the new html as appropriate
   - remove `_outl:___` tags from raindrop.io entries