import mammoth
from unidecode import unidecode
import glob, os
from datetime import datetime
import codecs

from rich.console import Console
from markdownify import markdownify

console = Console()

# Removes filetype extension
def filetype_remover(file):
    # Converts file name into a usable string
    current_file = str(file)
    # Separates the string into three components
    head, sep, tail = current_file.partition(".")  
    # Prints the filename with filetype removed  
    console.print("Filename:     " + head, style="#0015ff")
    # Returns the filename without filetype
    return(head)

# Function for .docx to .html
def docx_to_html_converter():
    # Loading the "docs" folder
    os.chdir("docs")
    # Loops through each .docx file in the folder
    for files in glob.glob("*.docx"):
        # Prints the current docx file
        console.print("File:         " + files, style="#0015ff")
        # Removes filetype extension and returns the name
        file_name = filetype_remover(files)
        # Opening the current file
        doc = mammoth.convert_to_html(files)
        # Prints html of file
        console.print("HTML:        ",doc.value, style="#EE6026")
        # Opens file in write mode
        f = open(file_name + '.html', "w")
        f.write(unidecode(doc.value))
        # Closes file
        f.close()

# Function for .html to .mk
def html_to_md_converter():
    # Loops through each .html file
    for files in glob.glob("*.html"):
        # Prints the current html file
        console.print("File:         " + files, style="#0015ff")
        # Removes file extension
        file_name = filetype_remover(files)
        # Convert the .html file to .mk
        file = open(files, "r").read()
        doc = markdownify(file, heading_style="ATX")
        # Prints the .mk
        console.print("MK:          ",doc, style="#EE6026")
        # Opens file in write mode
        f = open(file_name + '.md', "w")
        f.write(doc)
        # Closes file
        f.close()

# Function for docx to md
def docx_to_mk():
    print()
    console.print("-------------------------------------------", style="red")
    console.print("Converting .docx files to .html............", style="bold blue")
    console.print("-------------------------------------------", style="red")
    print()
    # Calls the docx to html function
    docx_to_html_converter()
    print()
    console.print("-------------------------------------------", style="red")
    console.print("Successfully converted .docx files to .html", style="bold green")
    console.print("-------------------------------------------", style="red")
    print()
    console.print("-------------------------------------------", style="red")
    console.print("Converting .html file to .md...............", style="blue")
    console.print("-------------------------------------------", style="red")
    print()
    # Calls the html to md function
    html_to_md_converter()
    print()
    console.print("-----------------------------------------", style="red")
    console.print("Successfully converted .html files to .md", style="bold green")
    console.print("-----------------------------------------", style="red")

# Run function
def run():
    convert_q = str(input("Do you want to convert files from .docx to .mk? Y/n:     "))
    if convert_q == "Y" or convert_q == "y":
        # Calls conversion functions
        docx_to_mk()
    else:
        console.print("Program aborted :(", style="bold red")

# Starts code
run()