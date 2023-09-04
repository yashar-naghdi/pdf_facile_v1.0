# The list of libraries needed are introduced here
import tkinter as tk
import fitz  # PyMuPDF
from tkinter import filedialog, ttk, PanedWindow , Frame
from utils import extract_annotations, save_to_excel, clear_markings
import tkinter.filedialog as filedialog
from tkinter import messagebox
import pandas as pd

# Initial an empty dataframe
df = pd.DataFrame()


# Defining global variables here
current_page = 0
rect = None
doc = None
start_x = None
start_y = None
pdf_page_size = (0, 0)
marked_areas = {}
last_selected_header= None
data_for_excel = {}
all_rectangles = {}

# The body of the code containing all the functions, starts here
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
    render_page(current_page)  # First, render the page at the new zoom level
    adjust_annotations_for_zoom(1.3)  # Next, adjust the annotations

def zoom_out():
    global zoom_level
    if zoom_level > 0.3:
        zoom_level -= 0.3
        render_page(current_page)  # First, render the page at the new zoom level
        adjust_annotations_for_zoom(0.9)  # Next, adjust the annotations


def adjust_annotations_for_zoom(scale_factor):
    if current_page not in all_rectangles:
        return  # No annotations to adjust on this page
    
    new_rectangles = []

    for rectangle in all_rectangles[current_page]:
        x0, y0, x1, y1 = canvas.coords(rectangle)
        
        # Adjust coordinates based on zoom level
        x0 *= scale_factor
        y0 *= scale_factor
        x1 *= scale_factor
        y1 *= scale_factor

        # Delete the old rectangle and create a new one with adjusted coordinates
        canvas.delete(rectangle)
        new_rectangle = canvas.create_rectangle(x0, y0, x1, y1, fill="green")
        new_rectangles.append(new_rectangle)
    
    all_rectangles[current_page] = new_rectangles

def next_page():
    global current_page
    total_pages = len(doc)
    if current_page < total_pages - 1:
        current_page +=1
        render_page(current_page)

def previous_page():
    global current_page
    if current_page > 0:
        current_page -= 1
        render_page(current_page)

def clear_all_markings():
    global marked_areas, data_for_excel, last_selected_header
    marked_areas = {}
    data_for_excel = {}
    last_selected_header = None

    # Clear the canvas markings
    canvas.delete('marking')

    # Clear the Excel preview (Treeview)
    clear_treeview(tree)

    
#  This function will clear all rows from the Treeview.
def clear_treeview(tree):
    for row in tree.get_children():
        tree.delete(row)


def handle_save():
    if not data_for_excel or all(len(data) == 0 for data in data_for_excel.values()):
        tk.messagebox.showinfo("Info", "No data to save.")
        return
    
    # Convert the data to a pandas DataFrame
    df = pd.DataFrame(data_for_excel)
    
    save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    if not save_path:
        return
    
    df.to_excel(save_path, index=False)
    tk.messagebox.showinfo("Success", f"Data saved to {save_path}")


def render_page(page_number):
    global pdf_page_size, current_page, doc, marked_areas

    # Clear the canvas
    canvas.delete("all")

    pdf_page = doc[page_number]
    pdf_page_size = (pdf_page.rect.width, pdf_page.rect.height)
    img = pdf_page.get_pixmap(matrix=fitz.Matrix(zoom_level, zoom_level))
    img.save("temp.png")
    pdf_image = tk.PhotoImage(file="temp.png")
    canvas.create_image(0, 0, anchor=tk.NW, image=pdf_image)
    canvas.image = pdf_image

    # Call this AFTER rendering the page to ensure the rectangles are on top.
    redraw_marked_areas()

def redraw_marked_areas():
    for page_num, areas in marked_areas.items():
        if page_num == current_page:
            for area in areas:
                x0, y0, x1, y1 = area
                x0 *= zoom_level
                y0 *= zoom_level
                x1 *= zoom_level
                y1 *= zoom_level
                canvas.create_rectangle(x0, y0, x1, y1, outline="green", tag='marking')



def on_canvas_click(event):
    global start_x, start_y, rect
    start_x, start_y = canvas.canvasx(event.x), canvas.canvasy(event.y)
    if rect:
        canvas.delete(rect)
    rect = canvas.create_rectangle(start_x, start_y, start_x, start_y, outline='red', width=4, stipple='gray50', tag="dragging")

def on_canvas_drag(event):
    global rect
    current_x, current_y = canvas.canvasx(event.x), canvas.canvasy(event.y)
    if rect:
        canvas.delete(rect)
    rect = canvas.create_rectangle(start_x, start_y, current_x, current_y, outline='red', width=4, stipple='gray50', tag="dragging")
#start_x, start_y = None, None






def on_canvas_release(event):
    global start_x, start_y, rect, end_x, end_y, doc, zoom_level, marked_areas, last_selected_header
    
    end_x, end_y = canvas.canvasx(event.x), canvas.canvasy(event.y)

    # Convert to original PDF dimensions for extraction
    pdf_start_x, pdf_start_y = start_x / zoom_level, start_y / zoom_level
    pdf_end_x, pdf_end_y = end_x / zoom_level, end_y / zoom_level

 # Store rectangle in marked_areas
    if current_page not in marked_areas:
        marked_areas[current_page] = []
    marked_areas[current_page].append((pdf_start_x, pdf_start_y, pdf_end_x, pdf_end_y))

    # Make sure start is always the top-left corner
    if pdf_start_x > pdf_end_x:
        pdf_start_x, pdf_end_x = pdf_end_x, pdf_start_x
    if pdf_start_y > pdf_end_y:
        pdf_start_y, pdf_end_y = pdf_end_y, pdf_start_y
    
    page = doc[current_page]
    rect = (pdf_start_x, pdf_start_y, pdf_end_x, pdf_end_y)
    extracted_text = page.get_text("text", clip=rect)

    extracted_data = extract_annotations(doc, marked_areas)
    formatted_data = [list(row.values()) for row in extracted_data]  # Convert each dict to a list of values
    update_treeview(formatted_data)


# The core of our code: 
    if event.state & 0x1:  # Shift key held during drag
        last_selected_header = extracted_text.strip()
        data_for_excel[last_selected_header] = []
    else:
        if last_selected_header:  # Ensure a header was previously selected
            data_for_excel[last_selected_header].append(extracted_text.strip())
        else:
            print("Please choose a header first")

    # Visualize the selection and then immediately delete the marking
    #visualize_selection((start_x, start_y, end_x, end_y))
    
    # Remove any existing red dragging rectangle and green marking rectangle
    canvas.delete('dragging')
    
    green_rect = visualize_selection((start_x, start_y, end_x, end_y))
    #canvas.after(3000, lambda: canvas.delete(green_rect))  # Delete after 3 seconds
    start_x, start_y = None, None







def extract_text_from_area(start, end):
    global current_page, doc
    scale = pdf_page_size[0] / canvas.winfo_width()
    rect = fitz.Rect(start[0] * scale, start[1] * scale, end[0] * scale, end[1] * scale)
    page = doc[current_page]
    text = page.get_text("text", clip=rect)
    return text
# Needs to be filled with the Excel Logic
def add_header_to_excel_preview(header_text):
    global df


def visualize_selection(rect_coords):
    x0, y0, x1, y1 = rect_coords
    canvas.create_rectangle(x0, y0, x1, y1, outline="green", width=2, tag='marking')


def on_key_press(event):
    global shift_pressed
    if event.keysym == 'Shift_L' or event.keysym == 'Shift_R':
        shift_pressed = True

def on_key_release(event):
    global shift_pressed
    if event.keysym == 'Shift_L' or event.keysym == 'Shift_R':
        shift_pressed = False

def on_mouse_wheel(event):
    if event.delta > 0:
        zoom_in()
    else:
        zoom_out()

def save_to_excel(filename):
    df.to_excel(filename, index=False)

# This function will be responsible for updating the Treeview (Excel preview) based on the PDF annotations.
def update_treeview(data):
    # Clear the treeview
    for row in tree.get_children():
        tree.delete(row)
    
    # Populate the treeview with the new data
    for row_data in data:
        tree.insert('', 'end', values=row_data)
# All the design aspects of the main application window and UI components goes here

root = tk.Tk()
root.title("pdf_facile")
root.geometry("1000x1000")

# Overall container for pane and side_frame
container = tk.Frame(root)
container.pack(fill=tk.BOTH, expand=True)

pane = tk.PanedWindow(container, orient=tk.HORIZONTAL)
pane.pack(side=tk.LEFT, fill=tk.BOTH, expand=1)

# Left side - Excel-like preview
excel_frame = tk.Frame(pane)
num_columns = 6
columns = [f"Col{i}" for i in range(1, num_columns + 1)]
tree = ttk.Treeview(excel_frame, columns=columns, show='headings')

for col in columns:
    tree.heading(col, text=col)
    tree.column(col, width=45) 

tree.pack(fill=tk.BOTH, expand=1)
pane.add(excel_frame)

# Right side - PDF Frame and components
pdf_outer_frame = tk.Frame(pane)
pdf_outer_frame.pack(side=tk.LEFT, padx=20, pady=20, fill=tk.BOTH, expand=True)
pdf_frame = tk.Frame(pdf_outer_frame, bg="light blue", width=595, height=842)  # Size for A4 page
pdf_frame.pack(fill=tk.BOTH, expand=True)
pane.add(pdf_outer_frame)

canvas = tk.Canvas(pdf_frame, bg="white", width=595, height=842)  # Size for A4 page
canvas.pack(fill=tk.BOTH, expand=True)
# Navigation Frame below the PDF Canvas
nav_frame = tk.Frame(pdf_outer_frame, bg="light blue")
nav_frame.pack(pady=20)


#setting standard button size
standard_button_below = 20
standard_button_side = 10

# Event bindings are described here
# Buttons section

previous_button = tk.Button(nav_frame, text="Previous", command=previous_page, width=standard_button_below)
previous_button.grid(row=0, column=0, padx=5, pady=5)

zoom_out_button = tk.Button(nav_frame, text="Zoom Out", command=zoom_out, width=standard_button_below)
zoom_out_button.grid(row=1, column=0, padx=5, pady=5)

next_button = tk.Button(nav_frame, text="Next", command=next_page, width=standard_button_below)
next_button.grid(row=0, column=1, padx=5, pady=5)

zoom_in_button = tk.Button(nav_frame, text="Zoom In", command=zoom_in, width=standard_button_below)
zoom_in_button.grid(row=1, column=1, padx=5, pady=5)

side_frame = tk.Frame(container, bg="light blue", width=400)
side_frame.pack(side=tk.RIGHT, padx=20, pady=20, fill=tk.Y)

pdf_button = tk.Button(side_frame, text="Open PDF",  bg="black", fg="white", command=open_pdf, width=standard_button_side)
pdf_button.pack(pady=20)

save_button = tk.Button(side_frame, text="Save",  bg="black", fg="white",command=handle_save, width= standard_button_side)
save_button.pack(pady=20)  # Adjust row and column as needed


clear_all_button = tk.Button(side_frame, text="Clear All",bg="black", fg="white", command=clear_all_markings, width=standard_button_side)
clear_all_button.pack(pady=20)

exit_button = tk.Button(side_frame, text="Exit",  bg="black", fg="white", command=root.quit, width=standard_button_side)
exit_button.pack(pady=20)

# Mouse section

canvas.bind("<ButtonPress-1>", on_canvas_click)     # Bind left mouse button click
canvas.bind("<B1-Motion>", on_canvas_drag)  # Bind dragging with left mouse button held down
canvas.bind("<ButtonRelease-1>", on_canvas_release) # Bind release of left mouse button
canvas.bind("<MouseWheel>", on_mouse_wheel)
canvas.bind("<KeyPress>", on_key_press)
canvas.bind("<KeyRelease>", on_key_release)
canvas.focus_set()

root.mainloop()
