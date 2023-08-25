import tkinter as tk
import fitz  # PyMuPDF
from tkinter import filedialog

current_page = 0
doc = None

def on_button_click():
    print("Button clicked!")

def open_pdf():
    global doc
    pdf_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
    
    if not pdf_path:
        return
    
    doc = fitz.open(pdf_path)
    render_page(0)

zoom_level = 1.0

def zoom_in():
    global zoom_level
    zoom_level += 0.3
    render_page(current_page)

def zoom_out():
    global zoom_level
    if zoom_level > 0.3:
        zoom_level -= 0.3
        render_page(current_page)

def render_page(page_number):
    global current_page, doc
    current_page = page_number
    page = doc[page_number]
    img = page.get_pixmap(matrix=fitz.Matrix(zoom_level, zoom_level))
    img.save("temp.png")
    pdf_image = tk.PhotoImage(file="temp.png")
    canvas.create_image(0, 0, anchor=tk.NW, image=pdf_image)
    canvas.image = pdf_image

def next_page():
    global current_page
    if current_page < len(doc) - 1:
        render_page(current_page + 1)

def previous_page():
    global current_page
    if current_page > 0:
        render_page(current_page - 1)

root = tk.Tk()
root.title("pdf_facile")
root.geometry("1000x1000")

pdf_frame = tk.Frame(root, bg="light blue")
pdf_frame.pack(side=tk.LEFT, padx=20, pady=20, fill=tk.BOTH, expand=True)

canvas = tk.Canvas(pdf_frame, bg="light blue")
canvas.pack(pady=20, fill=tk.BOTH, expand=True)

nav_frame = tk.Frame(pdf_frame, bg="light blue")
nav_frame.pack(pady=10)

#setting standard button size
standard_button_width = 20

previous_button = tk.Button(nav_frame, text="Previous", command=previous_page, width=standard_button_width)
previous_button.grid(row=0, column=0, padx=5, pady=5)

zoom_out_button = tk.Button(nav_frame, text="Zoom Out", command=zoom_out, width=standard_button_width)
zoom_out_button.grid(row=1, column=0, padx=5, pady=5)

next_button = tk.Button(nav_frame, text="Next", command=next_page, width=standard_button_width)
next_button.grid(row=0, column=1, padx=5, pady=5)

zoom_in_button = tk.Button(nav_frame, text="Zoom In", command=zoom_in, width=standard_button_width)
zoom_in_button.grid(row=1, column=1, padx=5, pady=5)

side_frame = tk.Frame(root, bg="gray", width=400)
side_frame.pack(side=tk.RIGHT, padx=20, pady=20, fill=tk.BOTH)

pdf_button = tk.Button(side_frame, text="Open PDF", command=open_pdf, width=20)
pdf_button.pack(pady=20)

button_ellipse = tk.Button(side_frame, text="Button1", bg="black", fg="white", command=on_button_click)
button_ellipse.pack(pady=20)

exit_button = tk.Button(side_frame, text="Exit", command=root.quit, width=20)
exit_button.pack(pady=20)

root.mainloop()
