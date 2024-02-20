# The list of libraries needed are introduced here
import tkinter as tk
import fitz  # PyMuPDF
from tkinter import filedialog, ttk, PanedWindow , Frame
# Assuming the utils module is available and contains the necessary functions
#import pillow as PIL
from PIL import Image, ImageTk
import tkinter.filedialog as filedialog
from tkinter import messagebox
import pandas as pd
import os
#import sys
import re
import io


# Setting the Class skeleton for the application
class PDFApp:
    def __init__(self, root):

        self.root = root
        self.current_page = 0
        self.current_column = 0
        self.column_indices = {}
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
        self.zoom_level = 1.0  # 1.0 means 100%, i.e., no zoom.
        self.last_marked_item = []
        self.green_rectangles = []
        self.data_entries = []
        self.line_counts = []
        self.marking_steps = []
        self.marked_info = []
        self.num_columns=10
        
        self.tree = None
        self.headers = [""] * self.num_columns  # We use this to remember the header of each column
        self.shift_pressed = False
        self.marked_areas = {}
        self.data_for_excel = {}
        self.data = {}
        self.last_selected_header = None
        self.grid_visible = False
        self.df = pd.DataFrame()
        self.setup_ui()
        self.rows = [self.tree.insert("", "end", values=("",) * self.num_columns) for _ in range(10)]
        self.row_index = 1  # Initialize row_index if you are using it
        self.placeholders = [self.tree.insert("", "end", values=("",) * self.num_columns) for _ in range(self.num_columns)]
        
        
    
    def on_canvas_click(self, event):
        
        self.start_x, self.start_y = self.canvas.canvasx(event.x), self.canvas.canvasy(event.y)
        if self.rect:
            self.canvas.delete(self.rect)
        self.rect = self.canvas.create_rectangle(self.start_x, self.start_y, self.start_x, self.start_y, outline='red', width=4, stipple='gray50', tag="dragging")

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
            x0, y0, x1, y1 = self.canvas.coords(rectangle)
            
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
        self.marked_areas = {}
        self.data_for_excel = {}
        self.last_selected_header = None
        self.current_page = 0  # or whatever your initial page is
        self.current_column = 0  # reset to the first column
        self.headers = [None] * self.num_columns  # reset headers
        self.column_indices = {}  # reset column indices
        # Clear the canvas markings
        self.canvas.delete('marking')

        # Clear the Excel preview (Treeview)
        self.clear_treeview(self.tree)

        # Reset the Treeview placeholders
        self.reset_treeview_placeholders()

    def clear_treeview(self, tree):
        for item in tree.get_children():
            tree.delete(item)

    def reset_treeview_placeholders(self):
        # Clear existing placeholders
        for placeholder in self.placeholders:
            if self.tree.exists(placeholder):
                self.tree.delete(placeholder)
        # Create new placeholders
        self.placeholders = [self.tree.insert("", "end", values=("",) * self.num_columns) for _ in range(1)]  # Only one placeholder row for headers

    def draw_grid(self):
        # Clear any existing grid lines
        self.canvas.delete("grid_line")
    
        # Define the spacing for the grid (e.g., every 50 pixels)
        grid_spacing_y = 20* self.zoom_level  # Adjust for zoom
        grid_spacing_x = 80* self.zoom_level  # Adjust for zoom
        
        # Define the color with reduced opacity (in this case, light green with 50% opacity)
        grid_color = '#D1F5D1'  # Light green
        extension_factor = 4
        # Draw vertical lines
        for i in range(0, int(self.canvas['width'])*extension_factor, int(grid_spacing_x)):
            self.canvas.create_line(i, 0, i, int(self.canvas['height'])*extension_factor,fill=grid_color, tag="grid_line")

        # Draw horizontal lines
        for i in range(0, int(self.canvas['height'])*extension_factor, int(grid_spacing_y)):
            self.canvas.create_line(0, i, int(self.canvas['width'])*extension_factor, i,fill=grid_color, tag="grid_line")

    def toggle_grid(self):
        if self.grid_visible:
            self.canvas.delete("grid_line")
            self.grid_visible = False
        else:
            self.draw_grid()
            self.grid_visible = True

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
        tk.messagebox.showinfo("Success", f"Data saved in {save_path}")

    def render_page(self,page_number):
        try:
            # Clear the canvas
            self.canvas.delete("all")

            pdf_page = self.doc[page_number]
            self.pdf_page_size = (pdf_page.rect.width, pdf_page.rect.height)
            img = pdf_page.get_pixmap(matrix=fitz.Matrix(self.zoom_level, self.zoom_level))
            img_data = img.tobytes("png")

            # Convert to a format Tkinter can use
            image = Image.open(io.BytesIO(img_data))
            photo = ImageTk.PhotoImage(image)

            self.canvas.create_image(0, 0, anchor=tk.NW, image= photo)
            self.canvas.image =  photo  # Keep a reference to avoid garbage collection

            # Call this AFTER rendering the page to ensure the rectangles are on top.
            self.redraw_marked_areas()

        except Exception as e:
            print(f"Error rendering page: {e}")

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

    def on_canvas_drag(self, event):
        
        current_x, current_y = self.canvas.canvasx(event.x), self.canvas.canvasy(event.y)
        if self.rect:
            self.canvas.delete(self.rect)
        self.rect = self.canvas.create_rectangle(self.start_x, self.start_y, current_x, current_y, outline='red', width=4, stipple='gray50', tag="dragging")

    def transform_decimal_separator(self, text):
        # This regular expression matches numbers with a period as the decimal separator
        pattern = r'(\d+\.\d+)'
        transformed_text = re.sub(pattern, lambda m: m.group(1).replace('.', ','), text)
        return transformed_text

    def format_name(self, name):
    # Remove any titles
        name_without_title = re.sub(r'(?:\b|^)(Monsieur|Mme\.|Madame|M\.|Mademoiselle)\s+', '', name).strip()
    
        # Split the name into words
        parts = name_without_title.split()
        
        if all(part.isupper() for part in parts) or all(not part.isupper() for part in parts):
            # For names not clearly distinguishable by case, assume last part is the last name
            last_name_parts = [parts[-1]]
            first_name_parts = parts[:-1]
        else:
            # Identify the last name (all CAPS) and first name
            last_name_parts = [part for part in parts if part.isupper()]
            first_name_parts  = [part for part in parts if part not in last_name_parts]
        
        # If we couldn't identify parts clearly, return the original
        if not last_name_parts  or not first_name_parts:
            return name.upper()
        
        # Ensure last name is first and everything is uppercase
        formatted_name = ' '.join(last_name_parts + first_name_parts).upper()
        return formatted_name
    
    def format_number(self,data):
        formatted_data = []
        for entry in data:
            # Remove spaces used for thousands
            entry_no_spaces = entry.replace(" ", "")
            # Replace comma with dot for decimals
            entry_standardized_decimal = entry_no_spaces.replace(",", ".")
            
            # Check if the entry is a number after replacements
            if re.match(r'^-?\d+(\.\d+)?$', entry_standardized_decimal):
                # Convert to appropriate numeric type
                if "." in entry_standardized_decimal:
                    formatted_data.append(float(entry_standardized_decimal))
                else:
                    formatted_data.append(int(entry_standardized_decimal))
            else:
                formatted_data.append(entry)  # Keep as is if not a number
        return formatted_data

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
        formatted_lines = self.format_number(extracted_lines)
        data_entry = [self.last_selected_header] + formatted_lines
        
        # Check if the extracted text is intended to be a name (i.e., it's in the first column)
        is_name_column = self.current_column == 0

        # If extracting names, format them
        if is_name_column:
            formatted_lines = [self.format_name(line) for line in formatted_lines]

        self.data_entries.append(data_entry)
        self.line_counts.append(len(formatted_lines))

            
        # This following section is the most delicate and the heart of our application
        
        # If shift is held, set the last selected line as a new header
        if event.state & 0x1:  # Shift key held during drag (Choosing a header)
            extracted_header = ' '.join(extracted_lines)
            #print(f"Extracted Header: {extracted_header}")
            # Update last_selected_header to the newly extracted header
            self.last_selected_header = extracted_header
            self.current_column = self.headers.index(extracted_header) if extracted_header in self.headers else self.current_column
            
        # Only update if it's a new header, otherwise continue using the current column
            if extracted_header not in self.headers:

                #print(f"Corrected Placeholder Index for {extracted_header}: {placeholder_index}")
                
                self.headers[self.current_column] = extracted_header
                
                self.column_indices[extracted_header] = 0  # Initialize the index for this column
                              
                placeholder_index = self.current_column
                
                # Update the header in the Treeview
                # Ensure there are enough placeholders
                while len(self.placeholders) <= self.current_column:
                    new_placeholder = self.tree.insert("", "end", values=("",) * self.num_columns)
                    self.placeholders.append(new_placeholder)

                self.tree.item(self.placeholders[self.current_column], values=("",) * self.current_column + (extracted_header,) + ("",) * (self.num_columns - self.current_column - 1))
                
        else:  # Data selection
            header = self.headers[self.current_column] if self.last_selected_header is not None else self.headers[self.current_column] # Get the header for the current column
            
            #print(f"Current Header: {header}")
            # If there's a header set for the column, insert the data under it
            if header:
                if header not in self.data_for_excel:
                    self.data_for_excel[header] = []
                if header not in self.column_indices:
                    self.column_indices[header] = 0
                self.data_for_excel[header].extend(formatted_lines)
                
                # Inserting data under the header in the Treeview
                for line in formatted_lines:
                    #print(f"Data Index for {header}: {self.column_indices[header]}")
                    self.tree.insert("", self.column_indices[header]+1 , values=("",) * self.current_column + (line,) + ("",) * (self.num_columns - self.current_column - 1))
                    self.column_indices[header] += 1  # Update the index for this column
                    #added some print statements to check the data

                #print(f"Data in Treeview after adding data: {[self.tree.item(child)['values'] for child in self.tree.get_children('')]}")

            # After inserting data, prepare for a new header (move to the next column)
            if self.current_column < self.num_columns - 1:
                self.current_column += 1 

        # Visualize the selection and remove existing rectangles
        self.canvas.delete('dragging')
        green_rect = self.visualize_selection((self.start_x, self.start_y, self.end_x, self.end_y))
        self.green_rectangles.append(green_rect)
        self.update_treeview(self.data_for_excel)

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
                    self.zoom_in()
        else:
                    self.zoom_out()

    def save_to_excel(self, filename):
        self.df.to_excel(filename, index=False)

    def scroll_left(self, event=None ):
        self.canvas.xview_scroll(-1, "units")

    def scroll_right(self, event=None):
        self.canvas.xview_scroll(1, "units")

    def move_up(self,event=None):
        self.canvas.yview_scroll(-1, "units")

    def move_down(self,event=None):
        self.canvas.yview_scroll(1, "units")

    # This function will be responsible for updating the Treeview (Excel preview) based on the PDF annotations.
    def update_treeview(self, data_entries):
        # Clear the treeview
        for item in self.tree.get_children():
            if item not in self.placeholders:
                #print(f"Deleting Item: {item}")
                self.tree.delete(item)
        
        # Find max length to align the columns
        max_len = max(len(v) for v in data_entries.values()) if data_entries else 0
        
        # Extract headers and transpose data for Treeview
        headers = list(data_entries.keys())
        
        transposed_data = []
        for i in range(max_len):
            row = []
            for header in headers:
                if i < len(data_entries[header]):
                    row.append(data_entries[header][i])
                else:
                    row.append("")  # Padding for shorter columns
            transposed_data.append(row)
        
        # Insert the new rows
        for entry in transposed_data:
            #print(f"Inserting Row: {entry}")
            self.tree.insert("", "end", values=entry)
    

    def setup_ui(self):
        # Overall container for pane and side_frame
        self.container = tk.Frame(self.root)
        self.container.pack(fill=tk.BOTH, expand=True)

        # Adding a scrollbar to the container
        self.scrollbar = tk.Scrollbar(self.container, orient='vertical')
        self.scrollbar.pack(side='right', fill='y')

        self.pane = tk.PanedWindow(self.container, orient=tk.HORIZONTAL)
        self.pane.pack(side=tk.LEFT, fill=tk.BOTH, expand=1)

        # Left side - Excel-like preview
        self.excel_frame = tk.Frame(self.pane)
        num_columns = 7
        columns = [f"Col{i}" for i in range(1, num_columns + 1)]
        self.tree = ttk.Treeview(self.excel_frame, columns=columns, show='headings')

        # Change the heading of the first column to "Name"
        self.tree.heading("Col1", text="Names")

        for col in columns[1:]:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=60) 

        self.tree.pack(fill=tk.BOTH, expand=1)
        self.pane.add(self.excel_frame)

        # Right side - PDF Frame and components
        self.pdf_outer_frame = tk.Frame(self.pane)
        self.pdf_outer_frame.pack(side=tk.LEFT, padx=20, pady=20, fill=tk.BOTH, expand=True)
        self.pdf_frame = tk.Frame(self.pdf_outer_frame, bg="light blue", width=595, height=842)  # Size for A4 page
        self.pdf_frame.pack(fill=tk.BOTH, expand=True)
        self.pane.add(self.pdf_outer_frame)

        # Create placeholder items for each column in the Treeview
        self.placeholders = [self.tree.insert("", "end", values=("",) * self.num_columns) for _ in range(1)]  # Only one placeholder row for headers
        #print(f"Placeholders: {self.placeholders}")

        #self.canvas = tk.Canvas(self.pdf_frame, bg="white", width=595, height=842)  # Size for A4 page
        self.canvas = tk.Canvas(self.pdf_frame, bg="white", width=595, height=650, yscrollcommand=self.scrollbar.set)  # Size for A4 page
        self.canvas.pack(fill=tk.BOTH, expand=True)
        self.scrollbar.config(command=self.canvas.yview)

        # Event bindings for canvas
        self.canvas.bind("<ButtonPress-1>", self.on_canvas_click)
        self.canvas.bind("<B1-Motion>", self.on_canvas_drag)
        self.canvas.bind("<ButtonRelease-1>", self.on_canvas_release)
        self.canvas.bind("<MouseWheel>", self.on_mouse_wheel)
        self.canvas.bind("<KeyPress>", self.on_key_press)
        self.canvas.bind("<KeyRelease>", self.on_key_release)
        self.canvas.bind("<Left>", self.scroll_left)
        self.canvas.bind("<Right>", self.scroll_right)
        self.canvas.bind("<Up>", self.move_up)
        self.canvas.bind("<Down>", self.move_down)
        
        self.canvas.bind("<KeyPress-z>", self.move_up)  # Equivalent to Up Arrow
        self.canvas.bind("<KeyPress-s>", self.move_down)  # Equivalent to Down Arrow
        self.canvas.bind("<KeyPress-q>", self.scroll_left)  # Equivalent to Left Arrow
        self.canvas.bind("<KeyPress-d>", self.scroll_right)  # Equivalent to Right Arrow
        
        self.canvas.focus_set()

        # Navigation Frame below the PDF Canvas
        self.nav_frame = tk.Frame(self.pdf_outer_frame, bg="#6ee2f5")
        self.nav_frame.pack(pady=20)

        #screen_width = self.root.winfo_screenwidth()
        
        #setting standard button size
        standard_button_below = 20
        standard_button_side = 10
        # Buttons section
        self.previous_button = tk.Button(self.nav_frame, text="Previous Page",  bg="black", fg="white",command=self.previous_page, width=standard_button_below)
        self.previous_button.grid(row=0, column=0, padx=5, pady=5)

        self.zoom_out_button = tk.Button(self.nav_frame, text="Zoom Back", bg="black", fg="white", command=self.zoom_out, width=standard_button_below)
        self.zoom_out_button.grid(row=1, column=0, padx=5, pady=5)

        self.next_button = tk.Button(self.nav_frame, text="Next Page", bg="black", fg="white",command=self.next_page, width=standard_button_below)
        self.next_button.grid(row=0, column=1, padx=5, pady=5)

        self.zoom_in_button = tk.Button(self.nav_frame, text="Zoom In", bg="black", fg="white", command=self.zoom_in, width=standard_button_below)
        self.zoom_in_button.grid(row=1, column=1, padx=5, pady=5)

        self.side_frame = tk.Frame(self.container, bg="#6ee2f5", width=500)
        self.side_frame.pack(side=tk.RIGHT, padx=20, pady=20, fill=tk.Y)

        self.pdf_button = tk.Button(self.side_frame, text="Open", bg="black", fg="white", command=self.open_pdf, width=standard_button_side)
        self.pdf_button.pack(pady=20)

        self.save_button = tk.Button(self.side_frame, text="Save", bg="black", fg="white", command=self.handle_save, width=standard_button_side)
        self.save_button.pack(pady=20)

        self.clear_all_button = tk.Button(self.side_frame, text="Delete", bg="black", fg="white", command=self.clear_all_markings, width=standard_button_side)
        self.clear_all_button.pack(pady=20)

        self.grid_button = tk.Button(self.side_frame, text="Grid", bg="black", fg="white", command=self.toggle_grid,width=standard_button_side)
        self.grid_button.pack(pady=20)

        self.exit_button = tk.Button(self.side_frame, text="Close", bg="black", fg="white", command=self.root.quit, width=standard_button_side)
        self.exit_button.pack(pady=20)

        # Get the current script directory
        script_dir = os.path.dirname(os.path.abspath(__file__))
        # Create a path to the icon relative to the script
        icon_path = os.path.join(script_dir, 'start.ico')

        self.root.iconbitmap(icon_path)
        self.root.title("PDF Facile")
        
    def show_startup_image(self):
        # Create a new top-level window
        startup_window = tk.Toplevel(self.root)
        # Set the window size and position to be the same as the main app
        window_width = 600  # Width of your main app window
        window_height = 600  # Height of your main app window
        # Get screen width and height
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        # Calculate position x and y coordinates
        x = (screen_width/2) - (window_width/2)
        y = (screen_height/2) - (window_height/2)

        startup_window.geometry(f"{window_width}x{window_height}+{int(x)}+{int(y)}")
        startup_window.overrideredirect(True)  # Remove window decorations


        # Get the current script directory
        script_dir = os.path.dirname(os.path.abspath(__file__))
        # Create a path to the image relative to the script
        image_path = os.path.join(script_dir, 'start2.png')

        # Load and show the image
        image = Image.open(image_path)
        photo = ImageTk.PhotoImage(image)
        label = tk.Label(startup_window, image=photo)
        label.image = photo  # Keep a reference to avoid garbage collection
        label.pack()

        # Close the startup window after 1000 milliseconds (1 second)
        startup_window.after(2500, startup_window.destroy)


root = tk.Tk()
app = PDFApp(root)
# Show the startup image
app.show_startup_image()
root.mainloop()
