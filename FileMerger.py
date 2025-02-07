import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from PyPDF2 import PdfMerger
from PIL import Image
from fpdf import FPDF
import os
import subprocess


class FileMerger:
    def __init__(self, root):
        self.standard_format = {
            ".pdf": tk.BooleanVar(value=True)
        }
        self.other_formats = {
            ".txt": tk.BooleanVar(value=False),
        }

        self.image_types = {".png", ".jpg", ".jpeg"}
        for ext in self.image_types:
            self.other_formats[ext] = tk.BooleanVar(value=False)

        self.root = root
        self.root.title("DocMerger")
        self.root.geometry("500x550")

        # Variables
        self.source_dir = ""
        self.doc_vars = {}
        self.selected_format = tk.StringVar(value="pdf")
        self.extensions_filter = {".pdf"}  # Default filter to PDFs

        # Styling
        s = ttk.Style()
        s.configure('TButton', padding=3)
        s.configure('TLabelframe', padding=5)
        s.configure('FileRow.TFrame', background='white')
        s.configure('Dragged.TFrame', background='white', relief='solid')
        s.configure('DragHandle.TLabel', font=('Segoe UI', 12), foreground='#999999')
        s.configure('DragLine.TSeparator', background='#0078d7')
        s.configure('DropTarget.TFrame', background='#f0f9ff')
        s.configure('FileCheckbutton.TCheckbutton', padding=5)
        s.configure('Files.TLabelframe', padding=10)

        # Main layout
        main = ttk.Frame(root, padding=5)
        main.pack(fill=tk.BOTH, expand=True)

        # Directory selection
        df = ttk.Frame(main)
        df.pack(fill=tk.X, pady=2)
        self.dir_entry = ttk.Entry(df)
        self.dir_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 2))
        ttk.Button(df, text="Browse", command=self.browse, width=8).pack(side=tk.RIGHT)

        # Extension Filter (Checkboxes)
        filter_frame = ttk.Frame(main)
        filter_frame.pack(fill=tk.X, pady=2)
        ttk.Label(filter_frame, text="Filter by Extensions:").pack(side=tk.LEFT)

        self.checkbuttons = {}  # Store checkbuttons for later updates
        combined_formats = {**self.standard_format, **self.other_formats}

        for ext, var in combined_formats.items():
            chk = ttk.Checkbutton(filter_frame, text=ext, variable=var, command=self.filter_files)
            chk.pack(side=tk.LEFT)
            self.checkbuttons[ext] = chk

        # Files list
        ff = ttk.LabelFrame(main, text="Files")
        ff.pack(fill=tk.BOTH, expand=True, pady=5)

        # Select All/None buttons
        select_buttons_frame = ttk.Frame(ff)
        select_buttons_frame.pack(fill=tk.X, padx=5, pady=2)
        ttk.Button(select_buttons_frame, text="Select All", command=self.select_all_files).pack(side=tk.LEFT, padx=2)
        ttk.Button(select_buttons_frame, text="Select None", command=self.select_none_files).pack(side=tk.LEFT, padx=2)

        # Scrollable frame
        self.canvas = tk.Canvas(ff)
        sb = ttk.Scrollbar(ff, orient="vertical", command=self.canvas.yview)
        self.file_frame = ttk.Frame(self.canvas, style='Content.TFrame')
        
        self.file_frame.bind("<Configure>", 
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.create_window((0, 0), window=self.file_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=sb.set)

        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb.pack(side=tk.RIGHT, fill=tk.Y)

        # Format Switch (Initially Hidden)
        self.format_frame = ttk.Frame(main)
        self.format_frame.pack_forget()
        self.format_dropdown = ttk.Combobox(self.format_frame, textvariable=self.selected_format, state="readonly")
        self.format_dropdown.pack(side=tk.LEFT, padx=5)

        # Output
        of = ttk.Frame(main)
        of.pack(fill=tk.X, pady=2)
        ttk.Label(of, text="File Name:").pack(side=tk.LEFT)
        self.out_entry = ttk.Entry(of)
        self.out_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=2)

        # Button frame
        button_frame = ttk.Frame(main)
        button_frame.pack(pady=5)
        ttk.Button(button_frame, text="Merge Files", command=self.merge).pack(side=tk.LEFT, padx=2)
        ttk.Button(button_frame, text="Open Directory", command=self.open_result_directory).pack(side=tk.LEFT, padx=2)

        # Status
        self.status = ttk.Label(main, wraplength=400)
        self.status.pack()

    def browse(self):
        dir = filedialog.askdirectory()
        if dir:
            self.source_dir = dir
            self.dir_entry.delete(0, tk.END)
            self.dir_entry.insert(0, dir)
            self.update_list()

    def filter_files(self):
        """Handles the extension filter checkboxes and updates file list"""
        selected_filters = {ext for ext, var in {**self.standard_format, **self.other_formats}.items() if var.get()}
        self.extensions_filter = selected_filters
        self.update_list(preserve_selection=True)

    def on_file_select(self, *args):
        """Handle individual file selection changes"""
        # Check if any files are selected
        any_selected = any(var.get() for var in self.doc_vars.values())
        if any_selected:
            self.format_frame.pack(fill=tk.X, pady=2)
            self.update_format_options()
        else:
            self.format_frame.pack_forget()

    def update_list(self, preserve_selection=False):
        """Updates the file list based on the selected directory and extension filter"""
        # Store current selections if needed
        current_selections = {}
        if preserve_selection:
            current_selections = {file: var.get() for file, var in self.doc_vars.items()}

        # Clear existing variables
        self.doc_vars.clear()

        if self.source_dir:
            for f in sorted(os.listdir(self.source_dir)):
                ext = os.path.splitext(f)[1].lower()
                if ext in self.extensions_filter:
                    # Default to unselected (0), unless we're preserving a previous selected state
                    initial_value = current_selections.get(f, 0)
                    var = tk.IntVar(value=initial_value)
                    var.trace_add('write', self.on_file_select)  # Add trace to variable
                    self.doc_vars[f] = var

            # After creating all variables, refresh the display
            self.refresh_file_list()

        # Show or hide format options based on selection state
        self.on_file_select()

    def update_format_options(self):
        """Update the export options based on selected file types"""
        selected_files = [os.path.join(self.source_dir, file) for file, var in self.doc_vars.items() if var.get()]
        
        if not selected_files:
            self.format_frame.pack_forget()
            return
        
        file_types = {os.path.splitext(f)[1].lower() for f in selected_files}

        # Always allow PDF as default
        available_formats = [".pdf"]

        # If all selected files are text files, allow txt format
        if file_types == {".txt"}:
            available_formats.append(".txt")
        # If all selected files are images, allow png format
        elif all(ext in self.image_types for ext in file_types):
            available_formats.append(".png")

        self.format_dropdown.config(values=available_formats)

        if self.selected_format.get() not in available_formats:
            self.selected_format.set(".pdf")

    def merge(self):
        """Merges the selected files and exports them in the chosen format"""
        if not self.source_dir:
            messagebox.showerror("Error", "Select directory first!")
            return

        out_name = self.out_entry.get().strip()
        if not out_name:
            messagebox.showerror("Error", "Enter output filename!")
            return

        # Get selected files in the order they appear in self.doc_vars
        selected = [os.path.join(self.source_dir, file) for file, var in self.doc_vars.items() if var.get()]

        if not selected:
            messagebox.showerror("Error", "Select at least one file!")
            return

        result_dir = os.path.join(self.source_dir, "result")
        
        # Check if the directory exists
        if not os.path.exists(result_dir):
            os.makedirs(result_dir)

        out_file = os.path.normpath(os.path.join(result_dir, f"{out_name}{self.selected_format.get().lower()}"))

        # Check if file exists and ask for confirmation
        if os.path.exists(out_file):
            if not messagebox.askyesno("Confirm Overwrite", 
                f"The file '{os.path.basename(out_file)}' already exists.\nDo you want to overwrite it?"):
                return

        try:
            format_type = self.selected_format.get().lower()

            if format_type == ".png":
                # Get dimensions for the combined image
                images = []
                total_height = 0
                max_width = 0

                # Process images in the order they appear in the list
                for file in selected:
                    img = Image.open(file)
                    if img.mode != 'RGB':
                        img = img.convert('RGB')
                    images.append(img)
                    total_height += img.height
                    max_width = max(max_width, img.width)

                # Create a new image with the combined dimensions
                combined_image = Image.new('RGB', (max_width, total_height), 'white')
                
                # Paste each image in order
                y_offset = 0
                for img in images:
                    # Center the image if it's smaller than max_width
                    x_offset = (max_width - img.width) // 2
                    combined_image.paste(img, (x_offset, y_offset))
                    y_offset += img.height

                # Save the combined image
                combined_image.save(out_file, 'PNG')
                
                if os.path.exists(out_file):
                    self.status.config(text=f"✓ Success! Saved as: {out_file}")
                return

            elif format_type == ".txt":
                # Merge text files in order
                with open(out_file, 'w', encoding='utf-8') as outfile:
                    for idx, file in enumerate(selected):
                        with open(file, 'r', encoding='utf-8') as infile:
                            if idx > 0:  # Add a separator between files
                                outfile.write('\n\n' + '='*50 + '\n\n')
                            outfile.write(infile.read())
                
                if os.path.exists(out_file):
                    self.status.config(text=f"✓ Success! Saved as: {out_file}")
                return

            elif format_type == ".pdf":
                pdf_merger = PdfMerger()
                temp_pdf_files = []

                # Convert each file to PDF in order
                for file in selected:
                    if file.lower().endswith(".pdf"):
                        pdf_merger.append(file)
                    elif file.lower().endswith((".jpg", ".jpeg", ".png")):
                        img = Image.open(file).convert("RGB")
                        temp_pdf_path = os.path.join(result_dir, f"temp_{os.path.basename(file)}.pdf")
                        img.save(temp_pdf_path)
                        pdf_merger.append(temp_pdf_path)
                        temp_pdf_files.append(temp_pdf_path)
                    elif file.lower().endswith(".txt"):
                        pdf = FPDF()
                        pdf.add_page()
                        pdf.set_auto_page_break(auto=True, margin=15)
                        pdf.set_font("Arial", size=12)
                        with open(file, "r", encoding="utf-8") as f:
                            text_content = f.read()
                            pdf.multi_cell(0, 10, text_content)
                        temp_pdf_path = os.path.join(result_dir, f"temp_{os.path.basename(file)}.pdf")
                        pdf.output(temp_pdf_path)
                        pdf_merger.append(temp_pdf_path)
                        temp_pdf_files.append(temp_pdf_path)

                if len(pdf_merger.pages) > 0:
                    with open(out_file, 'wb') as f_out:
                        pdf_merger.write(f_out)
                        pdf_merger.close()
                
                # Clean up temporary files
                for temp_pdf in temp_pdf_files:
                    os.remove(temp_pdf)

            else:
                self.status.config(text="NotImplementedException")

            if os.path.exists(out_file):
                self.status.config(text=f"✓ Success! Saved as: {out_file}")

        except Exception as e:
            self.status.config(text=f"Error: {str(e)}")
            print(f"Merge failed: {e}")

    def open_result_directory(self):
        """Opens the result directory in file explorer"""
        result_dir = os.path.join(self.source_dir, "result")
        if os.path.exists(result_dir):
            if os.name == 'nt':  # For Windows
                os.startfile(result_dir)
            else:  # For Linux/Mac
                subprocess.run(['xdg-open' if os.name == 'posix' else 'open', result_dir])
        else:
            messagebox.showinfo("Info", "Result directory does not exist yet. Merge some files first!")

    def select_all_files(self):
        """Select all files in the list"""
        for var in self.doc_vars.values():
            var.set(1)
        if self.doc_vars:
            self.format_frame.pack(fill=tk.X, pady=2)
            self.update_format_options()
        else:
            self.format_frame.pack_forget()

    def select_none_files(self):
        """Deselect all files in the list"""
        for var in self.doc_vars.values():
            var.set(0)
        self.format_frame.pack_forget()
        self.update_format_options()

    def start_drag(self, event, row_frame):
        """Start dragging a file"""
        row_frame.drag_start_y = event.y_root
        
        # Find the checkbutton in the content frame and get its text
        for child in row_frame.winfo_children():
            if isinstance(child, ttk.Frame):  # content frame
                for content_child in child.winfo_children():
                    if isinstance(content_child, ttk.Checkbutton):
                        row_frame.dragged_file = content_child.cget('text')
                        break
        
        if not hasattr(row_frame, 'dragged_file'):
            return  # Exit if we couldn't find the file name
            
        # Create a drag indicator line
        self.drag_line = ttk.Separator(self.file_frame, orient='horizontal', style='DragLine.TSeparator')
        # Configure drag style and lift
        row_frame.configure(style='Dragged.TFrame')
        row_frame.lift()
        
        # Add shadow effect
        row_frame.configure(relief='solid', borderwidth=1)

    def drag(self, event, row_frame):
        """Handle dragging of a file"""
        if not hasattr(row_frame, 'drag_start_y') or not hasattr(row_frame, 'dragged_file'):
            return

        # Get mouse position relative to file frame
        mouse_y = event.y_root - self.file_frame.winfo_rooty()
        
        # Get all row frames except the dragged one
        frames = [w for w in self.file_frame.winfo_children() 
                 if isinstance(w, ttk.Frame) and w != row_frame and not isinstance(w, ttk.Separator)]
        
        # Move the dragged frame with slight transparency effect
        row_frame.place(y=max(0, mouse_y - row_frame.winfo_height()/2))
        
        # Reset all frames to normal style
        for frame in frames:
            frame.configure(style='FileRow.TFrame')
        
        # Get current index of dragged file in the ordered dictionary
        files = list(self.doc_vars.keys())
        try:
            current_index = files.index(row_frame.dragged_file)
            
            # Find insertion point and show drag line
            new_index = current_index  # Default to current position
            insert_y = 0
            
            # If no frames, just show the line at mouse position
            if not frames:
                insert_y = mouse_y
            else:
                # Get the frame positions and heights
                frame_positions = [(frame.winfo_y(), frame.winfo_height()) for frame in frames]
                
                # Find the appropriate insertion point
                found_position = False
                for i, (frame_y, frame_height) in enumerate(frame_positions):
                    frame_bottom = frame_y + frame_height
                    
                    if mouse_y < frame_y + (frame_height * 0.5):  # Upper half of the frame
                        new_index = i
                        if i >= current_index:
                            new_index = i
                        else:
                            new_index = i
                        insert_y = frame_y
                        frames[i].configure(style='DropTarget.TFrame')
                        found_position = True
                        break
                
                # If we haven't found a position, we're at the bottom
                if not found_position:
                    new_index = len(frames)
                    if current_index < len(frames):  # Only adjust if we're moving down
                        new_index = len(frames)
                    insert_y = frame_positions[-1][0] + frame_positions[-1][1]
            
            # Update drag line position with more visible style
            self.drag_line.place(y=insert_y, relwidth=1, height=2)
            self.drag_line.lift()
            row_frame.lift()

            # Only reorder if the position has changed
            if new_index != current_index:
                # Move the file to new position
                files_copy = files.copy()
                file_to_move = files_copy.pop(current_index)
                files_copy.insert(new_index, file_to_move)
                
                # Recreate ordered dictionary with new order
                selections = {f: self.doc_vars[f].get() for f in self.doc_vars}
                self.doc_vars.clear()
                for f in files_copy:
                    var = tk.IntVar(value=selections[f])
                    var.trace_add('write', self.on_file_select)
                    self.doc_vars[f] = var
                
                # Update display but keep the dragged frame
                self.refresh_file_list(dragged_frame=row_frame)
                
                # Restore drag state
                row_frame.drag_start_y = event.y_root
                row_frame.dragged_file = file_to_move
                row_frame.lift()
                self.drag_line.lift()
        except ValueError:
            pass

    def stop_drag(self, event, row_frame):
        """Stop dragging a file"""
        if hasattr(row_frame, 'drag_start_y'):
            del row_frame.drag_start_y
        if hasattr(row_frame, 'dragged_file'):
            del row_frame.dragged_file
        # Remove drag line
        if hasattr(self, 'drag_line'):
            self.drag_line.destroy()
            del self.drag_line
        # Update display normally
        self.refresh_file_list()

    def refresh_file_list(self, dragged_frame=None):
        """Refresh the file list display"""
        # Store current frames' positions
        current_frames = {}
        for w in self.file_frame.winfo_children():
            if isinstance(w, ttk.Frame) and w != dragged_frame:
                # Find the checkbutton in the content frame
                for child in w.winfo_children():
                    if isinstance(child, ttk.Frame):  # This is the content frame
                        for content_child in child.winfo_children():
                            if isinstance(content_child, ttk.Checkbutton):
                                current_frames[content_child.cget('text')] = w.winfo_y()
                                break
        
        # Clear all frames except the dragged one
        for w in self.file_frame.winfo_children():
            if w != dragged_frame and not isinstance(w, ttk.Separator):
                w.destroy()

        # Configure row style
        style = ttk.Style()
        style.configure('FileRow.TFrame', background='white')

        for idx, (filename, var) in enumerate(self.doc_vars.items()):
            # Skip creating new frame for dragged item
            if dragged_frame and hasattr(dragged_frame, 'dragged_file') and dragged_frame.dragged_file == filename:
                continue

            # Create a frame for each file row
            row_frame = ttk.Frame(self.file_frame, style='FileRow.TFrame')
            row_frame.pack(fill=tk.X, padx=8, pady=1)

            # Create a container for the drag handle
            handle_frame = ttk.Frame(row_frame, style='FileRow.TFrame')
            handle_frame.pack(side=tk.LEFT, padx=(0, 5))

            # Add drag indicator with improved visibility
            drag_handle = ttk.Label(handle_frame, text=" : ", cursor="fleur", style='DragHandle.TLabel')
            drag_handle.pack(padx=(5, 0), pady=5)

            # Create a container for the checkbox and filename that allows text wrapping
            content_frame = ttk.Frame(row_frame, style='FileRow.TFrame')
            content_frame.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))

            # Checkbutton for file selection with improved layout and text wrapping
            check = ttk.Checkbutton(content_frame, text=filename, variable=var, 
                                  style='FileCheckbutton.TCheckbutton', width=63)  # Set width to allow wrapping
            check.pack(side=tk.LEFT, fill=tk.X, expand=True, anchor="w")
            
            # Make both the handle and its container draggable
            for widget in (drag_handle, handle_frame):
                widget.bind('<Button-1>', lambda e, rf=row_frame: self.start_drag(e, rf))
                widget.bind('<B1-Motion>', lambda e, rf=row_frame: self.drag(e, rf))
                widget.bind('<ButtonRelease-1>', lambda e, rf=row_frame: self.stop_drag(e, rf))

            # Try to maintain position if it existed before
            if filename in current_frames:
                row_frame.update_idletasks()
                self.canvas.yview_moveto(current_frames[filename] / self.file_frame.winfo_height())


if __name__ == "__main__":
    root = tk.Tk()
    app = FileMerger(root)
    root.mainloop()