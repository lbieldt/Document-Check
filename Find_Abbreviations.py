import re # Used to filter the text
from typing import Counter # Used to count the words
import docx2txt # Used to import docx
from docx import Document # Used to save the word docx (but can also read it)

# Used for GUI diagram
import tkinter as tk
from tkinter import filedialog
root = tk.Tk()
root.withdraw()

file_path = filedialog.askopenfilename() # Load file path with GUI popup
text = docx2txt.process(file_path) # Open connection to Word Document, read as txt
# text = docx2txt.process("C:\Python\Acronyms\Test.docx") #Static path to test



# Test abbreviation expression here: https://regex101.com/r/A1SJoe/16
rx = r"\b[A-Z](?=([&.]?))(?:\1[A-Z])+\b" # Looks for all the double Cap letters or special chars
s = text # Load text string
abbreviations =  [x.group() for x in re.finditer(rx, s)] #Save abbreviations in a list
abbr_sorted = sorted(abbreviations) #Sort list alphabetically

count_abbr = Counter(abbr_sorted) #Count all the unique items => Counter({'IPS': 14, 'HSS': 2, 'RCCS': 2, 'RB': 1})
unique_abbr = list(dict.fromkeys(abbr_sorted)) #List all the unique items => ['HSS', 'HSS', 'IPS', 'IPS', 'IPS', 'IPS', 'IPS', 'IPS', 'IPS', 'IPS', 'IPS', 

# Create a new word document
document = Document()
# Create a table with all the unique Abbreviation
table = document.add_table(rows=1, cols=2) # Create table with two columns
hdr_cells = table.rows[0].cells

hdr_cells[0].text = 'Abbreviation'
hdr_cells[1].text = 'Occurrence'
for abbr in unique_abbr:
    row_cells = table.add_row().cells
    row_cells[0].text = abbr
    row_cells[1].text = str(count_abbr[abbr])

table.style ="Table Grid"

# Display 
print(count_abbr)
print(unique_abbr)

file_path_save = filedialog.asksaveasfilename() #Load file path with GUI popup
document.save(file_path_save)