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
# UPDATE 2026-03-15
## Transitioned from diigo to raindrop.io
 - imported all diigo bookmarks into raindrop
 - use expand_redirects.py to correct 'flip.it' shortcuts to full urls
 - raindroptagger.py created to collect pub dates and word counts from raindrop.io exported csv files
 - to run, use `rdtagger` snippet in an ipynb cell
 - Once updated csv (with pub dates and wordcounts) is obtained, ~can run `add_raindrop_to_outline.py`~ (this was intended to automate categorization for outlines, but isn't working as well as the original - which was generated using poe.com)
# UPDATE 2026-02-14
 - export all new raindrop imports to .csv and run `raindroptagger.py` on that .csv file to get wordcounts and pub dates
   - Use `rdtagger` snippet for ease
     - set input & output filenames and heartbeat timer within the snippet code
 - Delete unsorted items from raindrop.io and re-upload with new records that now contain wordcounts and pub dates
- In the .csv file, delete records that are not to be included in the AI-related html file
- Add outline sections to tags in the .csv file
  - in the form `, _outl:IV-C` using appropriate outline designators
 - run `python "Python/add_outl_articles.py"` on the exported csv to include in the "Capstone AI articles.html"
 - copy the new html to local project folder and push to website
# UPDATE 2026-09-04
  - <span style="color: green;">Added automated outline assignment via semantic analysis to raindroptagger.py</span>
  - export 'Unsorted' raindrop folder to .csv and run `raindroptagger.py` to generate new .csv with pub dates, word counts, and outline assignments for 'ai'-tagged articles.
  - DOUBLE CHECK outline assignments
  - upload the new .csv into raindrop and move articles to proper folders
  -  - run `python "Python/add_outl_articles.py"` on the exported csv to include in the "Capstone AI articles.html"
 - copy the new html to local project folder and push to website
