# Importing necessary items

import fitz
import pandas as pd


# This function will go through all the marked areas on the PDF and extract the text.

def extract_annotations(doc, marked_areas):
    annotations = []

    for page_num, areas in marked_areas.items():
        page = doc[page_num]
        for area in areas:
            rect = fitz.Rect(area)
            text = page.get_text("text", clip=rect)
            annotations.append({
                "page": page_num,
                "coords": area,
                "text": text.strip()
            })

    return annotations
#This function will save the extracted annotations to an Excel file using the pandas library.
def save_to_excel(annotations, filename="annotations.xlsx"):
    df = pd.DataFrame(annotations)
    df.to_excel(filename, index=False)

#  A utility function that clears all the stored marked areas:

def clear_markings():
    global marked_areas
    marked_areas = []
