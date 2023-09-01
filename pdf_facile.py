# The list of libraries needed are introduced here
import tkinter as tk
import fitz  # PyMuPDF
from tkinter import filedialog


# Defining global variables here
current_page = 0
rect = None
doc = None
pdf_page_size = (0, 0)

# The body of the code containing all the functions
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
    global pdf_page_size, current_page, doc
    pdf_page = doc[page_number]
    pdf_page_size = (pdf_page.rect.width, pdf_page.rect.height)
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

def on_canvas_click(event):
    global start_x, start_y
    start_x, start_y = event.x, event.y
    canvas.create_rectangle(start_x, start_y, start_x, start_y, outline='red', tag='marking')

def on_canvas_drag(event):
    global start_x, start_y, rect
    canvas.delete(rect)
    
    # Adjust for zoom level
    adjusted_end_x = start_x + (event.x - start_x) / zoom_level
    adjusted_end_y = start_y + (event.y - start_y) / zoom_level
    
    rect = canvas.create_rectangle(start_x, start_y, adjusted_end_x, adjusted_end_y, outline="red")

def on_canvas_release(event):
    # At this point, you'd typically store the coordinates, but for now, we're just printing them.
    end_x, end_y = event.x, event.y
    print(f"Coordinates: {start_x}, {start_y}, {end_x}, {end_y}")

def extract_text_from_area(start, end):
    global current_page, doc

    # Scale the coordinates from canvas size to PDF size
    # Here, the assumption is that the PDF fits perfectly into the canvas
    # If there's any scaling applied to fit the PDF into the canvas, this needs to be considered
    scale = pdf_page_size[0] / canvas.winfo_width()

    # Convert the start and end coordinates to a PyMuPDF rectangle
    rect = fitz.Rect(start[0] * scale, start[1] * scale, end[0] * scale, end[1] * scale)

    # Extract text
    page = doc[current_page]
    text = page.get_text("text", clip=rect)

    return text

def visualize_selection(rect):
    """Visualizes the calculated selection."""
    # Extract the rectangle's coordinates
    x0, y0, x1, y1 = rect

    # Draw a semi-transparent rectangle for visualization
    canvas.create_rectangle(x0, y0, x1, y1, outline="green", dash=(4, 4))

def on_canvas_release(event):
    global start_x, start_y, rect, zoom_level

    end_x, end_y = canvas.canvasx(event.x), canvas.canvasy(event.y)
    
    if rect:
        canvas.delete(rect)
    
    # Calculate region based on zoom level and user selection
    doc_rect = fitz.Rect(start_x / zoom_level, start_y / zoom_level, end_x / zoom_level, end_y / zoom_level)

    # Extract text from this region
    page = doc.load_page(current_page)
    extracted_text = page.get_text("text", clip=doc_rect)
    print(extracted_text)
    
    # After extracting text, visualize the calculated selection
    visualize_selection((start_x, start_y, end_x, end_y))



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

pdf_button = tk.Button(side_frame, text="Open PDF",  bg="black", fg="white", command=open_pdf, width=20)
pdf_button.pack(pady=20)

idle_button = tk.Button(side_frame, text="Idle button",  bg="black", fg="white", command=on_button_click)
idle_button.pack(pady=20)

exit_button = tk.Button(side_frame, text="Exit",  bg="black", fg="white", command=root.quit, width=20)
exit_button.pack(pady=20)

# Mouse section

canvas.bind("<Button-1>", on_canvas_click)     # Bind left mouse button click
canvas.bind("<B1-Motion>", on_canvas_drag)     # Bind dragging with left mouse button held down
canvas.bind("<ButtonRelease-1>", on_canvas_release) # Bind release of left mouse button


root.mainloop()
