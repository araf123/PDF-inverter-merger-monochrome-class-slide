import tkinter as tk
from tkinter import ttk, filedialog, messagebox, Listbox, Scrollbar, Checkbutton, BooleanVar, StringVar, Entry
import subprocess
import os
import sys
import tempfile
import threading
import queue
import re
import shutil
import time
from pypdf import PdfReader, PdfWriter
from PIL import Image, ImageOps
import numpy as np

# --- Helper Function to find Ghostscript ---
def find_ghostscript_executable():
    """Finds the Ghostscript executable on the system."""
    gs_names = ["gswin64c.exe", "gswin32c.exe", "gs.exe"]
    # Check PATH first
    for name in gs_names:
        if shutil.which(name):
            return name
    # Check common installation directories
    program_files = [os.environ.get("ProgramFiles", ""), os.environ.get("ProgramFiles(x86)", "")]
    for pf in program_files:
        if pf:
            for root, _, files in os.walk(pf):
                for name in gs_names:
                    if name in files and 'gs' in root.lower() and 'bin' in root.lower():
                        return os.path.join(root, name)
    return None

GS_EXECUTABLE = find_ghostscript_executable()

# --- Main Application Class ---
class PdfToolApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF Processor Suite")
        self.root.geometry("600x700")
        self.root.minsize(550, 600)

        # Configure resizing behavior for the main window
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)

        self.task_queue = queue.Queue()
        self.file_list_data = []

        main_frame = tk.Frame(root)
        main_frame.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(1, weight=1) # Let the listbox frame expand

        # --- Frames ---
        action_frame = tk.Frame(main_frame)
        action_frame.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        action_frame.columnconfigure((0,1), weight=1)

        list_frame = tk.Frame(main_frame)
        list_frame.grid(row=1, column=0, sticky="nsew")
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(0, weight=1)

        editor_frame = tk.Frame(main_frame, relief=tk.GROOVE, borderwidth=2)
        editor_frame.grid(row=2, column=0, sticky="ew", pady=10)
        editor_frame.columnconfigure(1, weight=1)
        
        options_frame = tk.Frame(main_frame, relief=tk.GROOVE, borderwidth=2)
        options_frame.grid(row=3, column=0, sticky="ew")

        process_button_frame = tk.Frame(main_frame)
        process_button_frame.grid(row=4, column=0, sticky="ew", pady=(10, 0))
        process_button_frame.columnconfigure(0, weight=1)
        
        progress_frame = tk.Frame(main_frame)
        progress_frame.grid(row=5, column=0, sticky="ew", pady=(5,0))
        progress_frame.columnconfigure(0, weight=1)
        
        # --- File List & Buttons ---
        tk.Button(action_frame, text="1. Add PDFs", command=self.add_files).grid(row=0, column=0, sticky="ew", padx=(0,5))
        tk.Button(action_frame, text="Clear List", command=self.clear_list).grid(row=0, column=1, sticky="ew", padx=(5,0))

        self.listbox = Listbox(list_frame, selectmode=tk.SINGLE)
        self.listbox.grid(row=0, column=0, sticky="nsew")
        scrollbar = Scrollbar(list_frame, orient="vertical", command=self.listbox.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.listbox.config(yscrollcommand=scrollbar.set)
        self.listbox.bind('<<ListboxSelect>>', self.on_file_select)
        
        reorder_frame=tk.Frame(list_frame);
        reorder_frame.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(5,0))
        reorder_frame.columnconfigure((0,1), weight=1)
        tk.Button(reorder_frame, text="Move Up ↑", command=lambda: self.move_item('up')).grid(row=0, column=0, sticky="ew", padx=(0,5))
        tk.Button(reorder_frame, text="Move Down ↓", command=lambda: self.move_item('down')).grid(row=0, column=1, sticky="ew", padx=(5,0))

        # --- Page Editor ---
        tk.Label(editor_frame, text="Page Editor (for selected file)", font=("Helvetica", 10, "bold")).grid(row=0, column=0, columnspan=2, sticky="w", padx=5, pady=2)
        self.editor_info_label = tk.Label(editor_frame, text="Select a file to edit its pages.")
        self.editor_info_label.grid(row=1, column=0, columnspan=2, sticky="w", padx=5)
        
        range_frame=tk.Frame(editor_frame);range_frame.grid(row=2, column=0, columnspan=2, sticky="ew", pady=5, padx=5);range_frame.columnconfigure(1, weight=1)
        tk.Label(range_frame, text="Remove pages (e.g. 5, 8-12):").grid(row=0, column=0, sticky="w")
        self.page_range_var=StringVar(); self.page_range_entry=Entry(range_frame, textvariable=self.page_range_var); self.page_range_entry.grid(row=0, column=1, sticky="ew", padx=5)
        self.apply_range_button=tk.Button(range_frame,text="Apply",command=self.apply_page_range); self.apply_range_button.grid(row=0, column=2, padx=(0,5))
        self.reset_range_button=tk.Button(range_frame,text="Reset",command=self.reset_page_range); self.reset_range_button.grid(row=0, column=3)
        self.toggle_editor_widgets('disabled')
        
        # --- Global Options ---
        tk.Label(options_frame, text="Global Processing Options", font=("Helvetica", 10, "bold")).pack(anchor="w", padx=5, pady=2)
        layout_frame = tk.Frame(options_frame); layout_frame.pack(fill="x", padx=5, pady=3)
        tk.Label(layout_frame,text="Layout:").pack(side="left"); self.layout_var=StringVar(value="1"); ttk.Combobox(layout_frame,textvariable=self.layout_var,values=["1","2","3","4"],state="readonly",width=10).pack(side="left",padx=5); tk.Label(layout_frame,text="(slides per page)").pack(side="left")
        self.invert_var=BooleanVar(value=True); Checkbutton(options_frame, text="Invert Colors (for black background PDFs)", variable=self.invert_var).pack(anchor="w",padx=5)
        self.monochrome_var=BooleanVar(value=True); Checkbutton(options_frame, text="Apply Smart Monochrome Filter (Pure B&W)", variable=self.monochrome_var).pack(anchor="w",padx=5)

        # --- Process Button & Progress Bar ---
        self.process_button=tk.Button(process_button_frame,text="2. Process & Save PDF",bg="#4CAF50",fg="white",font=("Helvetica",12,"bold"),command=self.start_processing_thread); self.process_button.grid(row=0, column=0, sticky="ew", ipady=8)
        self.status_label = tk.Label(progress_frame, text="Select files to begin.", font=("Helvetica", 9), fg="gray"); self.status_label.grid(row=0, column=0, sticky="ew")
        self.time_label = tk.Label(progress_frame, text="", font=("Helvetica", 9), fg="gray"); self.time_label.grid(row=1, column=0, sticky="ew")
        self.progress_bar=ttk.Progressbar(progress_frame,mode='determinate'); self.progress_bar.grid(row=2, column=0, sticky="ew", pady=(5,0))

    def get_pages_to_process(self):
        """Parses the file list and returns a list of (path, page_num) tuples."""
        pages_to_process = []
        for file_data in self.file_list_data:
            path = file_data['path']
            try:
                reader = PdfReader(path)
                total_pages_in_file = len(reader.pages)
                
                pages_to_keep = set(range(1, total_pages_in_file + 1))
                range_spec = file_data['pages_to_remove']
                
                if range_spec and range_spec != 'none':
                    pages_to_remove = set()
                    for part in range_spec.split(','):
                        part = part.strip()
                        if not part: continue
                        if '-' in part:
                            start, end = part.split('-', 1)
                            start = int(start) if start else 1
                            end = int(end) if end else total_pages_in_file
                            pages_to_remove.update(range(start, end + 1))
                        else:
                            pages_to_remove.add(int(part))
                    pages_to_keep -= pages_to_remove

                for page_num in sorted(list(pages_to_keep)):
                    pages_to_process.append((path, page_num))
            except Exception as e:
                # In a real app, you might want to show this error to the user
                print(f"Could not read {os.path.basename(path)}: {e}")
                continue
        return pages_to_process

    def start_processing_thread(self):
        if not self.file_list_data: return messagebox.showwarning("No Files", "Please add one or more PDF files.")
        output_path = filedialog.asksaveasfilename(title="Save final processed PDF...", defaultextension=".pdf", filetypes=(("PDF Files", "*.pdf"),))
        if not output_path: return
        
        self.process_button.config(state="disabled", text="Processing...")
        self.toggle_editor_widgets('disabled')

        thread_args = {
            "output_path": output_path,
            "layout": self.layout_var.get(),
            "do_invert": self.invert_var.get(),
            "do_monochrome": self.monochrome_var.get(),
            "pages_to_process": self.get_pages_to_process(),
            "queue": self.task_queue,
        }
        threading.Thread(target=self.run_processing_in_thread, kwargs=thread_args, daemon=True).start()
        self.check_queue()
        
    def process_image_intelligently(self, img):
        """Applies inversion and a smart monochrome filter to a PIL Image."""
        if self.invert_var.get():
            if img.mode == 'RGBA': # Remove alpha channel
                bg = Image.new('RGB', img.size, (255, 255, 255))
                bg.paste(img, mask=img.getchannel('A'))
                img = bg
            img = ImageOps.invert(img.convert('RGB'))

        if self.monochrome_var.get():
            # Convert to numpy array for fast processing
            data = np.array(img)
            
            # Convert to HSV to detect "colored" (non-grayscale) areas
            hsv_data = np.array(img.convert('HSV'))
            saturation = hsv_data[:, :, 1]
            
            # Create a mask for colored regions (e.g., the blue boxes, now yellow)
            # A low saturation threshold filters out B&W and grays
            colored_mask = (saturation > 50) 
            
            # Convert to grayscale to create the base monochrome image
            gray_data = np.array(img.convert('L'))
            
            # Simple threshold for a high-contrast B&W image
            # Anything not almost pure white becomes black
            bw_data = (gray_data > 240) * 255 
            
            # --- The Smart Correction ---
            # 1. On the B&W image, make the colored areas completely black
            bw_data[colored_mask] = 0
            
            # 2. On top of those now-black areas, make the light text white again
            # We identify the light text from the original grayscale image
            light_text_mask = (gray_data > 128)
            bw_data[colored_mask & light_text_mask] = 255
            
            # Convert numpy array back to a PIL image in 'L' (grayscale) mode
            img = Image.fromarray(bw_data.astype(np.uint8)).convert('L')
            
        return img
    
    def run_processing_in_thread(self, output_path, layout, do_invert, do_monochrome, pages_to_process, queue):
        temp_dir = tempfile.mkdtemp()
        processed_images = []
        total_pages = len(pages_to_process)
        start_time = time.time()
        
        try:
            if total_pages == 0:
                raise ValueError("No pages were selected or found in the provided files.")
                
            # --- The New Page-by-Page Workflow ---
            for i, (pdf_path, page_num) in enumerate(pages_to_process):
                queue.put(('progress', (i, total_pages, start_time)))

                # Create a temporary PDF with just this single page
                single_page_pdf_path = os.path.join(temp_dir, 'single_page.pdf')
                with open(pdf_path, 'rb') as f:
                    reader = PdfReader(f)
                    writer = PdfWriter()
                    writer.add_page(reader.pages[page_num - 1])
                    with open(single_page_pdf_path, 'wb') as out_f:
                        writer.write(out_f)
                
                # Render only this single-page PDF to a temporary PNG
                output_png_path = os.path.join(temp_dir, 'page.png')
                gs_command = [
                    GS_EXECUTABLE,
                    '-dQUIET', '-dSAFER',
                    '-sDEVICE=png16m',
                    '-r200', # 200 DPI for good print quality
                    f'-o{output_png_path}',
                    single_page_pdf_path
                ]
                subprocess.run(gs_command, check=True, creationflags=subprocess.CREATE_NO_WINDOW)
                
                # Open the rendered image and process it
                with Image.open(output_png_path) as img:
                    processed_img = self.process_image_intelligently(img)
                    # Convert to RGB for universal PDF compatibility
                    processed_images.append(processed_img.convert('RGB'))

            if not processed_images:
                raise ValueError("Image processing failed to produce any pages.")
            
            # --- N-Up Layout Assembly ---
            queue.put(('status', "Step 2/2: Assembling final PDF..."))
            if layout == "1":
                processed_images[0].save(
                    output_path, "PDF", resolution=200.0, save_all=True,
                    append_images=processed_images[1:]
                )
            else:
                final_pdf_single_pages = tempfile.NamedTemporaryFile(suffix=".pdf", delete=False).name
                processed_images[0].save(
                    final_pdf_single_pages, "PDF", resolution=200.0, save_all=True,
                    append_images=processed_images[1:]
                )
                
                # Use pypdf for reliable n-up layout
                n_up_layout(final_pdf_single_pages, output_path, int(layout))
                os.remove(final_pdf_single_pages)

            queue.put(('success', f"PDF successfully processed!\nSaved to: {output_path}"))
            
        except Exception as e:
            error_msg = f"An error occurred: {str(e)}"
            if isinstance(e, subprocess.CalledProcessError):
                error_output = e.stderr or e.stdout or 'No error output from Ghostscript.'
                error_msg = f"Ghostscript failed:\n\n{error_output}"
            queue.put(('error', error_msg))
        finally:
            shutil.rmtree(temp_dir, ignore_errors=True)

    def check_queue(self):
        try:
            msg_type, data = self.task_queue.get_nowait()
        except queue.Empty:
            self.root.after(100, self.check_queue)
            return

        if msg_type == 'progress':
            current, total, start_time = data
            progress_val = (current / total) * 100
            self.progress_bar['value'] = progress_val
            
            elapsed = time.time() - start_time
            if current > 0:
                eta = (elapsed / current) * (total - current)
                eta_str = f"ETA: {time.strftime('%M:%S', time.gmtime(eta))}"
            else:
                eta_str = "ETA: --:--"
            
            elapsed_str = f"Elapsed: {time.strftime('%M:%S', time.gmtime(elapsed))}"

            self.status_label.config(text=f"Step 1/2: Processing page {current + 1} of {total}...", fg="blue")
            self.time_label.config(text=f"{elapsed_str}, {eta_str}", fg="blue")
            self.check_queue()
        elif msg_type == 'status':
             self.status_label.config(text=data,fg="blue"); self.time_label.config(text="")
             self.check_queue()
        else:
            self.progress_bar['value'] = 0
            self.status_label.config(text="", fg="gray")
            self.time_label.config(text="", fg="gray")
            if msg_type == 'success':
                messagebox.showinfo("Success!", data)
                self.status_label.config(text="Done!", fg="darkgreen")
            else:
                messagebox.showerror("Error", data)
                self.status_label.config(text="An error occurred.", fg="red")
            
            self.process_button.config(state="normal", text="2. Process & Save PDF")
            if self.listbox.curselection():
                self.toggle_editor_widgets('normal')

    # --- UI Helper Functions (mostly unchanged) ---
    def on_file_select(self, event=None):
        if not (sel := self.listbox.curselection()): return
        idx=sel[0]; self.toggle_editor_widgets('normal'); file_data=self.file_list_data[idx]
        self.editor_info_label.config(text=f"Editing: {os.path.basename(file_data['path'])}");
        page_range_text = file_data['pages_to_remove'] if file_data['pages_to_remove']!='none' else '';
        self.page_range_var.set(page_range_text)

    def toggle_editor_widgets(self,state): [w.config(state=state)for w in[self.page_range_entry,self.apply_range_button,self.reset_range_button]]
    
    def add_files(self):
        files=filedialog.askopenfilenames(title="Add PDF Files",filetypes=(("PDF Files","*.pdf"),));
        if files:
            for f in files:
                if not any(d['path']==f for d in self.file_list_data):self.file_list_data.append({'path':f,'pages_to_remove':'none','display_name':f"{os.path.basename(f)} [All Pages]"})
            self.update_listbox(); self.status_label.config(text=f"{len(self.file_list_data)} file(s) in list.",fg="darkgreen")
            if not self.listbox.curselection(): self.listbox.selection_set(0); self.on_file_select()

    def apply_page_range(self):
        if not(sel_idx:=self.listbox.curselection()):return
        idx=sel_idx[0];range_spec=self.page_range_var.get().strip()
        if range_spec and not re.match(r'^[\d\s,-]+$',range_spec):return messagebox.showerror("Input Error","Invalid characters. Use numbers, commas, and hyphens.")
        self.file_list_data[idx]['pages_to_remove']=range_spec if range_spec else 'none'
        self.file_list_data[idx]['display_name']=f"{os.path.basename(self.file_list_data[idx]['path'])} [Removing: {range_spec}]" if range_spec else f"{os.path.basename(self.file_list_data[idx]['path'])} [All Pages]";self.update_listbox(selection_idx=idx)

    def reset_page_range(self):
        if not(sel_idx:=self.listbox.curselection()):return
        idx=sel_idx[0];self.file_list_data[idx]['pages_to_remove']='none';self.file_list_data[idx]['display_name']=f"{os.path.basename(self.file_list_data[idx]['path'])} [All Pages]";self.page_range_var.set('');self.update_listbox(selection_idx=idx)

    def update_listbox(self,selection_idx=None):
        self.listbox.delete(0,tk.END);[self.listbox.insert(tk.END,item['display_name'])for item in self.file_list_data];
        if selection_idx is not None: self.listbox.selection_set(selection_idx); self.listbox.see(selection_idx)

    def clear_list(self):
        self.file_list_data=[]; self.update_listbox(); self.status_label.config(text="Select files to begin.",fg="gray");
        self.toggle_editor_widgets('disabled'); self.editor_info_label.config(text="Select a file to edit its pages."); self.page_range_var.set("")
    
    def move_item(self,direction):
        if not(sel_idx:=self.listbox.curselection()):return
        pos=sel_idx[0];
        if direction=='up' and pos > 0: self.file_list_data.insert(pos-1,self.file_list_data.pop(pos)); self.update_listbox(pos-1)
        elif direction=='down' and pos < len(self.file_list_data)-1: self.file_list_data.insert(pos+1,self.file_list_data.pop(pos)); self.update_listbox(pos+1)

def n_up_layout(input_pdf_path, output_pdf_path, pages_per_sheet):
    """Creates an n-up layout PDF using pypdf."""
    reader = PdfReader(input_pdf_path)
    writer = PdfWriter()
    
    a4_w, a4_h = 595.2, 841.8  # A4 Portrait dimensions

    if pages_per_sheet == 2:
        # 2-up landscape
        new_page_width, new_page_height = a4_h, a4_w
        positions = [(0, 0), (a4_h / 2, 0)]
        slot_w, slot_h = a4_h / 2, a4_w
    elif pages_per_sheet == 3:
        # 3-up portrait
        new_page_width, new_page_height = a4_w, a4_h
        positions = [(0, a4_h * 2/3), (0, a4_h / 3), (0, 0)]
        slot_w, slot_h = a4_w, a4_h / 3
    elif pages_per_sheet == 4:
         # 4-up portrait
        new_page_width, new_page_height = a4_w, a4_h
        positions = [(0, a4_h * 3/4), (0, a4_h * 2/4), (0, a4_h/4), (0,0)]
        slot_w, slot_h = a4_w, a4_h/4
    else: # Fallback to 1-up
        shutil.copy(input_pdf_path, output_pdf_path)
        return

    for i in range(0, len(reader.pages), pages_per_sheet):
        new_page = writer.add_blank_page(width=new_page_width, height=new_page_height)
        page_chunk = reader.pages[i : i + pages_per_sheet]
        
        for j, page in enumerate(page_chunk):
            p_w, p_h = float(page.mediabox.width), float(page.mediabox.height)
            if p_w == 0 or p_h == 0: continue
            
            scale = min(slot_w / p_w, slot_h / p_h)
            tx = positions[j][0] + (slot_w - p_w * scale) / 2
            ty = positions[j][1] + (slot_h - p_h * scale) / 2
            
            # Use pypdf's Transformation and merge_page
            op = pypdf.Transformation().scale(scale).translate(tx, ty)
            new_page.merge_transformed_page(page, op)

    with open(output_pdf_path, "wb") as f:
        writer.write(f)

# --- Entry Point ---
if __name__ == "__main__":
    if GS_EXECUTABLE is None:
        messagebox.showerror("Critical Error", "Ghostscript not found! Please install Ghostscript and ensure it's in your system's PATH.\nDownload from ghostscript.com.")
    else:
        root = tk.Tk()
        app = PdfToolApp(root)
        root.mainloop()