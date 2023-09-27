# The list of libraries needed are introduced here
import tkinter as tk
import fitz  # PyMuPDF
from tkinter import filedialog, ttk, PanedWindow , Frame
from utils import extract_annotations, save_to_excel, clear_markings

from tkinter import messagebox
import pandas as pd
import itertools
import os
import sys


# Initializing an empty dataframe
df = pd.DataFrame()

# Setting the Class skeleton for the application
class PDFExtractor:
    def __init__(self, root):

        # Create main frames and other widgets
        self.side_frame = tk.Frame(self.root)
        self.side_frame.pack(side=tk.LEFT, fill=tk.BOTH)
        
        self.canvas = tk.Canvas(self.root, bg="white")
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        self.pdf_button = tk.Button(self.side_frame, text="Open PDF", command=self.open_pdf)
        self.pdf_button.pack(pady=20)

        self.current_page = 0
        self.current_column = 0
        self.rect = None
        self.doc = None
        self.start_x = None
        self.end_x = None
        self.end_y = None 
        self.start_y = None
        self.pdf_page_size = (0, 0)
        self.marked_areas = {}
        self.last_selected_header= None
        self.all_rectangles = {}
        self.zoom_level = 1.0
        self.last_marked_item = []
        self.green_rectangles = []
        self.data_entries = []
        self.line_counts = []
        self.marking_steps = []
        self.marked_info = []
        self.num_columns=10
        self.headers = [""] * self.num_columns  # We use this to remember the header of each column
        self.shift_pressed = False
        self.marked_areas = {}
        self.data_for_excel = {}
        self.last_selected_header = None
        
        
        self.canvas.bind("<ButtonPress-1>", self.on_canvas_click)     # Bind left mouse button click
        self.canvas.bind("<B1-Motion>", self.on_canvas_drag)  # Bind dragging with left mouse button held down
        self.canvas.bind("<ButtonRelease-1>", self.on_canvas_release) # Bind release of left mouse button
        self.canvas.bind("<MouseWheel>", self.on_mouse_wheel)
        self.anvas.bind("<KeyPress>", self.on_key_press)
        self.canvas.bind("<KeyRelease>", self.on_key_release)
        self.canvas.bind("<Left>", self.scroll_left)
        self.canvas.bind("<Right>", self.scroll_right)
        self.canvas.bind("<Up>", self.move_up)
        self.canvas.bind("<Down>", self.move_down)
        self.canvas.focus_set()

# The body of the code containing all the functions, starts here
def on_button_click(self):
    print("Button clicked!")

def open_pdf(self):
    
    pdf_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
    
    if not pdf_path:
        return
    
    self.doc = fitz.open(pdf_path)
    self.render_page(0)

def zoom_in(self):
    
    self.zoom_level += 0.3
    self.render_page(self.current_page)  # First, render the page at the new zoom level
    self.adjust_annotations_for_zoom(1.3)  # Next, adjust the annotations

def zoom_out(self):
    
    if self.zoom_level > 0.3:
        self.zoom_level -= 0.3
        self.render_page(self.current_page)  # First, render the page at the new zoom level
        self.adjust_annotations_for_zoom(0.9)  # Next, adjust the annotations


def adjust_annotations_for_zoom(self,scale_factor):
    if self.current_page not in self.all_rectangles:
        return  # No annotations to adjust on this page
    
    new_rectangles = []

    for rectangle in self.all_rectangles[self.current_page]:
        x0, y0, x1, y1 = canvas.coords(rectangle)
        
        # Adjust coordinates based on zoom level
        x0 *= scale_factor
        y0 *= scale_factor
        x1 *= scale_factor
        y1 *= scale_factor

        # Delete the old rectangle and create a new one with adjusted coordinates
        self.canvas.delete(rectangle)
        new_rectangle = self.canvas.create_rectangle(x0, y0, x1, y1, fill="green")
        new_rectangles.append(new_rectangle)
    
    self.all_rectangles[self.current_page] = new_rectangles

def next_page(self):
    
    total_pages = len(self.doc)
    if self.current_page < total_pages - 1:
        self.current_page +=1
        self.render_page(self.current_page)

def previous_page(self):
    
    if self.current_page > 0:
        self.current_page -= 1
        self.render_page(self.current_page)

def clear_treeview(self,tree):
    for row in tree.get_children():
        tree.delete(row)

def clear_all_markings(self):
   
    # Clear the canvas markings
    self.canvas.delete('marking')

    # Clear the Excel preview (Treeview)
    self.clear_treeview(self.tree)


def handle_save(self):
    # Check if there's data to save
    if not self.data_for_excel or all(len(data) == 0 for data in self.data_for_excel.values()):
        tk.messagebox.showinfo("Info", "No data to save.")
        return
    
    # Transpose the data to fit into a DataFrame.
    # Essentially, we're converting the dictionary to a list of lists, 
    # where each inner list represents a column in our excel file.
    data_list = list(self.data_for_excel.values())
    max_len = max(len(lst) for lst in data_list)
    
    for lst in data_list:
        lst.extend([''] * (max_len - len(lst)))
    
    # Convert the data to a pandas DataFrame
    df = pd.DataFrame(list(zip(*data_list)), columns=self.data_for_excel.keys())
    
    # Get path to save and save it
    save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    if not save_path:
        return
    df.to_excel(save_path, index=False)
    tk.messagebox.showinfo("Success", f"Data saved to {save_path}")



def render_page(self,page_number):
    
    # Clear the canvas
    canvas.delete("all")

    pdf_page = self.doc[page_number]
    self.pdf_page_size = (pdf_page.rect.width, pdf_page.rect.height)
    img = pdf_page.get_pixmap(matrix=fitz.Matrix(self.zoom_level, self.zoom_level))
    img.save("temp.png")
    pdf_image = tk.PhotoImage(file="temp.png")
    self.canvas.create_image(0, 0, anchor=tk.NW, image=pdf_image)
    self.canvas.image = pdf_image

    # Call this AFTER rendering the page to ensure the rectangles are on top.
    self.redraw_marked_areas()

    # Once loaded into the canvas, delete the file
    try:
        os.remove("temp.png")
    except Exception as e:
        print(f"Error removing temp.png: {e}")

def redraw_marked_areas(self):
    for page_num, areas in self.marked_areas.items():
        if page_num == self.current_page:
            for area in areas:
                x0, y0, x1, y1 = area
                x0 *= self.zoom_level
                y0 *= self.zoom_level
                x1 *= self.zoom_level
                y1 *= self.zoom_level
                self.canvas.create_rectangle(x0, y0, x1, y1, outline="green", tag='marking')


def on_canvas_click(self, event):
    
    self.start_x, self.start_y = self.canvas.canvasx(event.x), self.canvas.canvasy(event.y)
    if self.rect:
        self.canvas.delete(self.rect)
    self.rect = self.canvas.create_rectangle(self.start_x, self.start_y, self.start_x, self.start_y, outline='red', width=4, stipple='gray50', tag="dragging")

def on_canvas_drag(self, event):
    
    current_x, current_y = self.canvas.canvasx(event.x), self.canvas.canvasy(event.y)
    if self.rect:
        self.canvas.delete(self.rect)
    self.rect = self.canvas.create_rectangle(self.start_x, self.start_y, current_x, current_y, outline='red', width=4, stipple='gray50', tag="dragging")


def on_canvas_release(self,event):
    
    
    self.end_x, self.end_y = self.canvas.canvasx(event.x), self.canvas.canvasy(event.y)

    # Convert to original PDF dimensions for extraction
    pdf_start_x, pdf_start_y = self.start_x / self.zoom_level, self.start_y / self.zoom_level
    pdf_end_x, pdf_end_y = self.end_x / self.zoom_level, self.end_y / self.zoom_level

 # Store rectangle in marked_areas
    if self.current_page not in self.marked_areas:
        self.marked_areas[self.current_page] = []
    self.marked_areas[self.current_page].append((pdf_start_x, pdf_start_y, pdf_end_x, pdf_end_y))

    # Make sure start is always the top-left corner
    if pdf_start_x > pdf_end_x:
        pdf_start_x, pdf_end_x = pdf_end_x, pdf_start_x
    if pdf_start_y > pdf_end_y:
        pdf_start_y, pdf_end_y = pdf_end_y, pdf_start_y
    
    page = self.doc[self.current_page]
    self.rect = (pdf_start_x, pdf_start_y, pdf_end_x, pdf_end_y)
    extracted_text = page.get_text("text", clip=self.rect)
    
    extracted_lines = [line for line in extracted_text.strip().splitlines() if line.strip() != ""]
    data_entry = [self.last_selected_header] + extracted_lines
    self.data_entries.append(data_entry)
    self.line_counts.append(len(extracted_lines))
    # If shift is held, set the last selected line as a new header
    if event.state & 0x1:  # Shift key held during drag (Choosing a header)
        extracted_header = ' '.join(extracted_lines)

        # Only update if it's a new header, otherwise continue using the current column
        if extracted_header not in self.headers:
            self.headers[self.current_column] = extracted_header
            self.tree.insert("", "end", values=("",) * self.current_column + (extracted_header,) + ("",) * (self.num_columns - self.current_column - 1))
        else:
            self.current_column = self.headers.index(extracted_header)  # Use the column of the existing header

    else:  # Data selection
        header = self.headers[self.current_column]  # Get the header for the current column

        # If there's a header set for the column, insert the data under it
        if header:
            for line in extracted_lines:
                self.tree.insert("", "end", values=("",) * self.current_column + (line,) + ("",) * (self.num_columns - self.current_column - 1))
        else:
            print("Please choose a header first")

        # After inserting data, prepare for a new header (move to the next column)
        if self.current_column < self.num_columns - 1:
            self.current_column += 1
    
    

    # Visualize the selection and remove existing rectangles
    self.canvas.delete('dragging')
    green_rect = self.visualize_selection((self.start_x, self.start_y, self.end_x, self.end_y))
    self.green_rectangles.append(green_rect)
    



def extract_text_from_area(self, start, end):
    
    scale = self.pdf_page_size[0] / self.canvas.winfo_width()
    self.rect = fitz.Rect(start[0] * scale, start[1] * scale, end[0] * scale, end[1] * scale)
    page = self.doc[self.current_page]
    text = page.get_text("text", clip=self.rect)
    return text


def visualize_selection(self, rect_coords):
    x0, y0, x1, y1 = rect_coords
    self.canvas.create_rectangle(x0, y0, x1, y1, outline="green", width=2, tag='marking')


def on_key_press(self, event):
    
    if event.keysym == 'Shift_L' or event.keysym == 'Shift_R':
        self.shift_pressed = True

def on_key_release(self,event):
    
    if event.keysym == 'Shift_L' or event.keysym == 'Shift_R':
        self.shift_pressed = False

def on_mouse_wheel(self, event):
    if event.delta > 0:
        zoom_in()
    else:
        zoom_out()

def save_to_excel(self, filename):
    df.to_excel(filename, index=False)

def scroll_left(self, event):
    self.canvas.xview_scroll(-1, "units")

def scroll_right(self, event):
    self.canvas.xview_scroll(1, "units")

def move_up(self,event=None):
    self.canvas.yview_scroll(-1, "units")

def move_down(self,event=None):
    self.canvas.yview_scroll(1, "units")

# This function will be responsible for updating the Treeview (Excel preview) based on the PDF annotations.
def update_treeview(self,data_entries):
    self.tree.delete(*tree.get_children())  # Clear the treeview

    # Transpose the data_entries list for Treeview
    transposed_data = list(itertools.zip_longest(*data_entries, fillvalue=""))

    for entry in transposed_data:
        # Insert each row into Treeview
        values = entry + ("",) * (self.num_columns - len(entry))
        self.tree.insert("", "end", values=values)


 #All the design aspects of the main application window and UI components goes here

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

pdf_button = tk.Button(side_frame, text="Open PDF",  bg="black", fg="white", command=self.open_pdf, width=standard_button_side)
pdf_button.pack(pady=20)

save_button = tk.Button(side_frame, text="Save",  bg="black", fg="white",command=self.handle_save, width= standard_button_side)
save_button.pack(pady=20)  # Adjust row and column as needed


clear_all_button = tk.Button(side_frame, text="Clear All",bg="black", fg="white", command=clear_all_markings, width=standard_button_side)
clear_all_button.pack(pady=20)

#remove_last_btn = tk.Button(side_frame, text="Remove Last",bg="black", fg="white", command=remove_last_marking)
#remove_last_btn.pack(pady=10)

exit_button = tk.Button(side_frame, text="Exit",  bg="black", fg="white", command=root.quit, width=standard_button_side)
exit_button.pack(pady=20)



if __name__ == "__main__":
    root = tk.Tk()
    app = PDFExtractor(root)
    root.mainloop()


