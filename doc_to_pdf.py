# pip install comtypes
# pip install python-docx

import os
import comtypes.client
import docx

# Paths to your Word document and desired PDF file
word_path = "E:\Programming & Coding\Projects\Competitions and Hack a Thons\MINeD Hackathon\Submission\Chat PDF Huggingface\Experimental Inputs\Developer_team.docx"
pdf_path = "E:\Programming & Coding\Projects\Competitions and Hack a Thons\MINeD Hackathon\Submission\Chat PDF Huggingface\Experimental Inputs\Developer_team.pdf"

# Load the Word document
doc = docx.Document(word_path)

# Create a Word application object
word = comtypes.client.CreateObject("Word.Application")

# Get absolute paths for Word document and PDF file
docx_path = os.path.abspath(word_path)
pdf_path = os.path.abspath(pdf_path)

# PDF format code
pdf_format = 17

# Make Word application invisible
word.Visible = False

try:
    # Open the Word document
    in_file = word.Documents.Open(docx_path)

    # Save the document as PDF
    in_file.SaveAs(pdf_path, FileFormat=pdf_format)

    print("Conversion successful. PDF saved at:", pdf_path)

except Exception as e:
    print("Error:", e)

finally:
    # Close the Word document and quit Word application
    if 'in_file' in locals():
        in_file.Close()
    word.Quit()