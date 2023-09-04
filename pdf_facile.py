# The list of libraries needed are introduced here
import tkinter as tk
import fitz  # PyMuPDF
from tkinter import filedialog
from utils import extract_annotations, save_to_excel, clear_markings
import tkinter.filedialog as filedialog
from tkinter import messagebox

# Defining global variables here
current_page = 0
rect = None
doc = None
start_x = None
start_y = None
pdf_page_size = (0, 0)
marked_areas = {}

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
    # Scale all items (including rectangles) on the canvas
    canvas.scale("all", 0, 0, 1.3, 1.3)
    render_page(current_page)

def zoom_out():
    global zoom_level
    if zoom_level > 0.3:
        zoom_level -= 0.3
        # Scale all items (including rectangles) on the canvas
        canvas.scale("all", 0, 0, 0.9, 0.9)
        render_page(current_page)

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
    global marked_areas
    canvas.delete("marking")
    marked_areas.clear()


def handle_save():
    annotations = extract_annotations(doc, marked_areas)  # Using global variables
    if not annotations:
        tk.messagebox.showinfo("Info", "No annotations to save.")
        return
    save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    if not save_path:
        return
    save_to_excel(annotations, save_path)
    tk.messagebox.showinfo("Success", f"Annotations saved to {save_path}")


def render_page(page_number):
    global pdf_page_size, current_page, doc, marked_areas
    pdf_page = doc[page_number]
    pdf_page_size = (pdf_page.rect.width, pdf_page.rect.height)
    page = doc[page_number]
    img = page.get_pixmap(matrix=fitz.Matrix(zoom_level, zoom_level))
    img.save("temp.png")
    pdf_image = tk.PhotoImage(file="temp.png")
    canvas.create_image(0, 0, anchor=tk.NW, image=pdf_image)
    canvas.image = pdf_image
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
    global start_x, start_y
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
    global start_x, start_y, rect, doc, zoom_level, marked_areas
    
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

    if event.state & 0x1:  
        print(f"Column Header Selected: {extracted_text}")
    else:
        print(extracted_text)

    # Visualize the selection and then immediately delete the marking
    visualize_selection((start_x, start_y, end_x, end_y))
    
    # Remove any existing red dragging rectangle and green marking rectangle
    canvas.delete('dragging')
    canvas.delete('marking')
    green_rect = visualize_selection((start_x, start_y, end_x, end_y))
    canvas.after(3000, lambda: canvas.delete(green_rect))  # Delete after 3 seconds
    # Remember the marked area
    #original_coords = (start_x, start_y, end_x, end_y)
    #marked_areas.append(original_coords)

    start_x, start_y = None, None

def extract_text_from_area(start, end):
    global current_page, doc
    scale = pdf_page_size[0] / canvas.winfo_width()
    rect = fitz.Rect(start[0] * scale, start[1] * scale, end[0] * scale, end[1] * scale)
    page = doc[current_page]
    text = page.get_text("text", clip=rect)
    return text

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

# All the design aspects of the main application window and UI components goes here

root = tk.Tk()
root.title("pdf_facile")
root.geometry("1000x1000")

# Main PDF Frame
pdf_frame = tk.Frame(root, bg="light blue", width=595, height=842)  # Size for A4 page
pdf_frame.pack(side=tk.LEFT, padx=20, pady=20, fill=tk.BOTH, expand=True)

canvas = tk.Canvas(pdf_frame, bg="white", width=595, height=842)  # Size for A4 page
canvas.pack(fill=tk.BOTH, expand=True)  # Allow it to resize with window

nav_frame = tk.Frame(pdf_frame, bg="light blue")
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

side_frame = tk.Frame(root, bg="light blue", width=400)
side_frame.pack(side=tk.RIGHT, padx=20, pady=20, fill=tk.BOTH)

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
