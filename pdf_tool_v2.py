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
from pypdf import PdfReader, PdfWriter, Transformation
from PIL import Image, ImageOps
import numpy as np

# This function is correct and will work with the bundled GS directory
def find_ghostscript_executable():
    """
    Finds the Ghostscript executable, prioritizing one bundled in its own subdirectory.
    """
    gs_name = "gswin64c.exe"

    if getattr(sys, 'frozen', False):
        bundle_dir = sys._MEIPASS
        bundled_gs_path = os.path.join(bundle_dir, 'gs_bin', gs_name)
        if os.path.exists(bundled_gs_path):
            return bundled_gs_path

    dev_path = f"D:\\gs10.05.1\\bin\\{gs_name}"
    if os.path.exists(dev_path):
        return dev_path

    if shutil.which(gs_name):
        return shutil.which(gs_name)

    return None

GS_EXECUTABLE = find_ghostscript_executable()

# --- Main Application Class ---
class PdfToolApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF Processor Suite")
        self.root.geometry("600x700")
        self.root.minsize(550, 600)
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        self.task_queue = queue.Queue()
        self.file_list_data = []
        # Main frames and widgets setup is correct and unchanged...
        main_frame = tk.Frame(root); main_frame.grid(row=0, column=0, sticky="nsew", padx=10, pady=10); main_frame.columnconfigure(0, weight=1); main_frame.rowconfigure(1, weight=1)
        action_frame = tk.Frame(main_frame); action_frame.grid(row=0, column=0, sticky="ew", pady=(0, 10)); action_frame.columnconfigure((0,1), weight=1)
        list_frame = tk.Frame(main_frame); list_frame.grid(row=1, column=0, sticky="nsew"); list_frame.columnconfigure(0, weight=1); list_frame.rowconfigure(0, weight=1)
        editor_frame = tk.Frame(main_frame, relief=tk.GROOVE, borderwidth=2); editor_frame.grid(row=2, column=0, sticky="ew", pady=10); editor_frame.columnconfigure(1, weight=1)
        options_frame = tk.Frame(main_frame, relief=tk.GROOVE, borderwidth=2); options_frame.grid(row=3, column=0, sticky="ew")
        process_button_frame = tk.Frame(main_frame); process_button_frame.grid(row=4, column=0, sticky="ew", pady=(10, 0)); process_button_frame.columnconfigure(0, weight=1)
        progress_frame = tk.Frame(main_frame); progress_frame.grid(row=5, column=0, sticky="ew", pady=(5,0)); progress_frame.columnconfigure(0, weight=1)
        tk.Button(action_frame, text="1. Add PDFs", command=self.add_files).grid(row=0, column=0, sticky="ew", padx=(0,5))
        tk.Button(action_frame, text="Clear List", command=self.clear_list).grid(row=0, column=1, sticky="ew", padx=(5,0))
        self.listbox = Listbox(list_frame, selectmode=tk.SINGLE); self.listbox.grid(row=0, column=0, sticky="nsew")
        scrollbar = Scrollbar(list_frame, orient="vertical", command=self.listbox.yview); scrollbar.grid(row=0, column=1, sticky="ns"); self.listbox.config(yscrollcommand=scrollbar.set); self.listbox.bind('<<ListboxSelect>>', self.on_file_select)
        reorder_frame=tk.Frame(list_frame); reorder_frame.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(5,0)); reorder_frame.columnconfigure((0,1), weight=1)
        tk.Button(reorder_frame, text="Move Up ↑", command=lambda: self.move_item('up')).grid(row=0, column=0, sticky="ew", padx=(0,5))
        tk.Button(reorder_frame, text="Move Down ↓", command=lambda: self.move_item('down')).grid(row=0, column=1, sticky="ew", padx=(5,0))
        tk.Label(editor_frame, text="Page Editor (for selected file)", font=("Helvetica", 10, "bold")).grid(row=0, column=0, columnspan=2, sticky="w", padx=5, pady=2)
        self.editor_info_label = tk.Label(editor_frame, text="Select a file to edit its pages."); self.editor_info_label.grid(row=1, column=0, columnspan=2, sticky="w", padx=5)
        range_frame=tk.Frame(editor_frame); range_frame.grid(row=2, column=0, columnspan=2, sticky="ew", pady=5, padx=5); range_frame.columnconfigure(1, weight=1)
        tk.Label(range_frame, text="Remove pages (e.g. 5, 8-12):").grid(row=0, column=0, sticky="w")
        self.page_range_var=StringVar(); self.page_range_entry=Entry(range_frame, textvariable=self.page_range_var); self.page_range_entry.grid(row=0, column=1, sticky="ew", padx=5)
        self.apply_range_button=tk.Button(range_frame,text="Apply",command=self.apply_page_range); self.apply_range_button.grid(row=0, column=2, padx=(0,5))
        self.reset_range_button=tk.Button(range_frame,text="Reset",command=self.reset_page_range); self.reset_range_button.grid(row=0, column=3)
        self.toggle_editor_widgets('disabled')
        tk.Label(options_frame, text="Global Processing Options", font=("Helvetica", 10, "bold")).pack(anchor="w", padx=5, pady=2)
        layout_frame = tk.Frame(options_frame); layout_frame.pack(fill="x", padx=5, pady=3)
        tk.Label(layout_frame,text="Layout:").pack(side="left"); self.layout_var=StringVar(value="1"); ttk.Combobox(layout_frame,textvariable=self.layout_var,values=["1","2","3","4"],state="readonly",width=10).pack(side="left",padx=5); tk.Label(layout_frame,text="(pages per sheet)").pack(side="left")
        self.invert_var=BooleanVar(value=True); Checkbutton(options_frame, text="Invert Colors (for dark background PDFs)", variable=self.invert_var).pack(anchor="w",padx=5)
        self.monochrome_var=BooleanVar(value=True); Checkbutton(options_frame, text="Apply Smart Monochrome Filter (Pure B&W)", variable=self.monochrome_var).pack(anchor="w",padx=5)
        self.process_button=tk.Button(process_button_frame,text="2. Process & Save PDF",bg="#4CAF50",fg="white",font=("Helvetica",12,"bold"),command=self.start_processing_thread); self.process_button.grid(row=0, column=0, sticky="ew", ipady=8)
        self.status_label = tk.Label(progress_frame, text="Select files to begin.", font=("Helvetica", 9), fg="gray"); self.status_label.grid(row=0, column=0, sticky="ew")
        self.time_label = tk.Label(progress_frame, text="", font=("Helvetica", 9), fg="gray"); self.time_label.grid(row=1, column=0, sticky="ew")
        self.progress_bar=ttk.Progressbar(progress_frame,mode='determinate'); self.progress_bar.grid(row=2, column=0, sticky="ew", pady=(5,0))

    def get_pages_to_process(self):
        pages_to_process = []
        for file_data in self.file_list_data:
            path = file_data['path']
            try:
                with open(path, 'rb') as f:
                    reader = PdfReader(f)
                    total_pages_in_file = len(reader.pages)
                pages_to_keep = set(range(1, total_pages_in_file + 1))
                range_spec = file_data.get('pages_to_remove', 'none')
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
                        else: pages_to_remove.add(int(part))
                    pages_to_keep -= pages_to_remove
                for page_num in sorted(list(pages_to_keep)): pages_to_process.append((path, page_num))
            except Exception as e:
                self.task_queue.put(('error', f"Could not read {os.path.basename(path)}: {e}"))
                return []
        return pages_to_process

    def start_processing_thread(self):
        if not self.file_list_data: return messagebox.showwarning("No Files", "Please add one or more PDF files.")
        pages_to_process = self.get_pages_to_process()
        if not pages_to_process: self.check_queue(); return
        output_path = filedialog.asksaveasfilename(title="Save final processed PDF...", defaultextension=".pdf", filetypes=(("PDF Files", "*.pdf"),))
        if not output_path: return
        self.process_button.config(state="disabled", text="Processing..."); self.toggle_editor_widgets('disabled')
        threading.Thread(target=self.run_processing_in_thread, kwargs={"output_path": output_path, "layout": self.layout_var.get(), "do_invert": self.invert_var.get(), "do_monochrome": self.monochrome_var.get(), "pages_to_process": pages_to_process, "queue": self.task_queue}, daemon=True).start()
        self.check_queue()

    def process_image_intelligently(self, img, do_invert, do_monochrome):
        """[FINAL MEMORY-EFFICIENT VERSION]"""
        if do_invert:
            if img.mode == 'RGBA':
                bg = Image.new('RGB', img.size, (255, 255, 255))
                bg.paste(img, mask=img.getchannel('A'))
                img = bg
            else:
                img = img.convert('RGB')
            img = ImageOps.invert(img)
        
        if do_monochrome:
            # Create a final image canvas, initialized to all white. This is our output.
            final_image_data = np.full(img.size[::-1], 255, dtype=np.uint8)
            
            # Convert to grayscale and get the data. We only need this one array.
            gray_img = img.convert('L')
            gray_data = np.array(gray_img)
            
            # Condition 1: Basic high contrast. Anywhere the gray is not bright, make it black.
            final_image_data[gray_data < 220] = 0
            
            # Condition 2: Smart filter for colored boxes.
            hsv_img = img.convert('HSV')
            hsv_data = np.array(hsv_img)
            saturation = hsv_data[:, :, 1]
            
            # Find colored regions using a temporary boolean mask
            colored_mask = saturation > 40
            
            # In the colored regions, force pixels to be black...
            final_image_data[colored_mask] = 0
            # ...UNLESS the original pixel was bright (this is the white text).
            final_image_data[colored_mask & (gray_data > 150)] = 255
            
            # Clean up intermediate arrays explicitly to be safe
            del gray_data, hsv_data, saturation
            
            # Convert our final numpy canvas back to a PIL Image
            img = Image.fromarray(final_image_data)
            
        return img

    def run_processing_in_thread(self, output_path, layout, do_invert, do_monochrome, pages_to_process, queue):
        temp_dir = tempfile.mkdtemp()
        try:
            total_pages = len(pages_to_process)
            if total_pages == 0: raise ValueError("No pages were selected.")
            queue.put(('progress', (0, total_pages, time.time()))) 
            
            final_pdf_parts = []
            
            is_processing_needed = do_invert or do_monochrome
            if is_processing_needed:
                BATCH_SIZE = 20; batch_images = []; batch_counter = 0

                for i, (pdf_path, page_num) in enumerate(pages_to_process):
                    queue.put(('progress', (i, total_pages, time.time())))
                    with open(pdf_path, "rb") as f:
                        reader = PdfReader(f); page_to_process = reader.pages[page_num - 1]
                        single_page_pdf_path = os.path.join(temp_dir, 'single_page.pdf')
                        writer_single = PdfWriter(); writer_single.add_page(page_to_process)
                        with open(single_page_pdf_path, 'wb') as out_f: writer_single.write(out_f)

                    output_png_path = os.path.join(temp_dir, 'page.png')
                    subprocess.run([GS_EXECUTABLE, '-dQUIET', '-dSAFER', '-sDEVICE=png16m', '-r200', f'-o{output_png_path}', single_page_pdf_path], check=True, creationflags=subprocess.CREATE_NO_WINDOW)
                    with Image.open(output_png_path) as img:
                        processed_img = self.process_image_intelligently(img, do_invert, do_monochrome)
                        batch_images.append(processed_img.convert('RGB')) # Convert to RGB for saving
                    
                    if len(batch_images) >= BATCH_SIZE or (i + 1) == total_pages:
                        queue.put(('status', f"Step 1/3: Assembling batch {batch_counter + 1}..."))
                        if batch_images:
                            batch_filename = os.path.join(temp_dir, f'batch_{batch_counter}.pdf')
                            batch_images[0].save(batch_filename, "PDF", resolution=200.0, save_all=True, append_images=batch_images[1:])
                            final_pdf_parts.append(batch_filename)
                            batch_images.clear()
                            batch_counter += 1
            else:
                queue.put(('status', f"Step 1/3: Collecting pages..."))
                writer = PdfWriter()
                for pdf_path, page_num in pages_to_process:
                     with open(pdf_path, "rb") as f:
                        reader = PdfReader(f); writer.add_page(reader.pages[page_num - 1])
                unprocessed_pdf_path = os.path.join(temp_dir, "unprocessed.pdf")
                with open(unprocessed_pdf_path, "wb") as f: writer.write(f)
                final_pdf_parts.append(unprocessed_pdf_path)

            queue.put(('status', "Step 2/3: Merging all parts..."))
            merged_final_path = os.path.join(temp_dir, 'merged.pdf')
            if not final_pdf_parts: raise ValueError("No PDF parts created to merge.")
            elif len(final_pdf_parts) == 1: shutil.copy(final_pdf_parts[0], merged_final_path)
            else:
                gs_merge_cmd = [GS_EXECUTABLE, '-q', '-dNOPAUSE', '-dBATCH', '-sDEVICE=pdfwrite', f'-sOutputFile={merged_final_path}'] + final_pdf_parts
                subprocess.run(gs_merge_cmd, check=True, creationflags=subprocess.CREATE_NO_WINDOW)
            
            queue.put(('status', f"Step 3/3: Saving final '{layout}-up' layout..."))
            if layout == "1": shutil.copy(merged_final_path, output_path)
            else: n_up_layout(merged_final_path, output_path, int(layout))
            
            queue.put(('success', f"PDF successfully processed!\nSaved to: {output_path}"))
            
        except Exception as e:
            error_msg = f"An error occurred: {str(e)}";
            if isinstance(e, subprocess.CalledProcessError): error_output = e.stderr.decode(errors='ignore') if e.stderr else (e.stdout.decode(errors='ignore') if e.stdout else 'No error output.'); error_msg = f"Ghostscript failed:\n\n{error_output}"
            queue.put(('error', error_msg))
        finally: shutil.rmtree(temp_dir, ignore_errors=True)

    def check_queue(self):
        try: msg_type, data = self.task_queue.get_nowait()
        except queue.Empty: self.root.after(100, self.check_queue); return
        if msg_type == 'progress':
            current, total, start_time = data; progress_val = (current / total) * 100
            self.progress_bar['value'] = progress_val; elapsed = time.time() - start_time
            eta_str = f"ETA: {time.strftime('%M:%S', time.gmtime((elapsed / (current+1)) * (total - (current+1))))}" if current+1 < total else "ETA: 00:00"
            elapsed_str = f"Elapsed: {time.strftime('%M:%S', time.gmtime(elapsed))}"
            if "Assembling batch" in self.status_label.cget("text"): pass
            else: self.status_label.config(text=f"Processing page {current + 1} of {total}...", fg="blue")
            self.time_label.config(text=f"{elapsed_str}, {eta_str}", fg="blue"); self.root.after(100, self.check_queue)
        elif msg_type == 'status':
             self.status_label.config(text=data,fg="blue"); self.time_label.config(text=""); self.root.after(100, self.check_queue)
        else:
            self.progress_bar['value'] = 0; self.time_label.config(text="")
            if msg_type == 'success': messagebox.showinfo("Success!", data); self.status_label.config(text="Done!", fg="darkgreen")
            else: messagebox.showerror("Error", data); self.status_label.config(text="An error occurred.", fg="red")
            self.process_button.config(state="normal", text="2. Process & Save PDF")
            if self.listbox.curselection(): self.toggle_editor_widgets('normal')

    # UI Helper Functions (Unchanged)
    def on_file_select(self, event=None):
        if not (sel := self.listbox.curselection()): return
        idx=sel[0]; self.toggle_editor_widgets('normal'); file_data=self.file_list_data[idx]
        self.editor_info_label.config(text=f"Editing: {os.path.basename(file_data['path'])}"); page_range_text = file_data.get('pages_to_remove', 'none'); self.page_range_var.set("" if page_range_text == 'none' else page_range_text)
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
        if range_spec and not re.match(r'^[\d\s,-]+$',range_spec):return messagebox.showerror("Input Error","Invalid characters.")
        self.file_list_data[idx]['pages_to_remove']=range_spec if range_spec else 'none'; self.file_list_data[idx]['display_name']=f"{os.path.basename(self.file_list_data[idx]['path'])} [Removing: {range_spec}]" if range_spec else f"{os.path.basename(self.file_list_data[idx]['path'])} [All Pages]";self.update_listbox(selection_idx=idx)
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

# This function is correct.
def n_up_layout(input_pdf_path, output_pdf_path, pages_per_sheet):
    reader = PdfReader(input_pdf_path)
    writer = PdfWriter()
    a4_w, a4_h = 595.2, 841.8
    if pages_per_sheet == 2: new_page_width, new_page_height = a4_h, a4_w; positions = [(0, 0), (a4_h / 2, 0)]; slot_w, slot_h = a4_h / 2, a4_w
    elif pages_per_sheet == 3: new_page_width, new_page_height = a4_w, a4_h; positions = [(0, a4_h * 2/3), (0, a4_h / 3), (0, 0)]; slot_w, slot_h = a4_w, a4_h / 3
    elif pages_per_sheet == 4: new_page_width, new_page_height = a4_w, a4_h; positions = [(0, a4_h * 3/4), (0, a4_h * 2/4), (0, a4_h/4), (0,0)]; slot_w, slot_h = a4_w, a4_h/4
    else: raise ValueError(f"Unsupported layout: {pages_per_sheet}")

    for i in range(0, len(reader.pages), pages_per_sheet):
        new_page = writer.add_blank_page(width=new_page_width, height=new_page_height)
        for j, page in enumerate(reader.pages[i : i + pages_per_sheet]):
            p_w, p_h = page.mediabox.width, page.mediabox.height
            if p_w == 0 or p_h == 0: continue
            scale = min(slot_w / p_w, slot_h / p_h)
            tx = positions[j][0] + (slot_w - p_w * scale) / 2
            ty = positions[j][1] + (slot_h - p_h * scale) / 2
            page.add_transformation(Transformation().scale(scale).translate(tx, ty))
            new_page.merge_page(page)
    with open(output_pdf_path, "wb") as f: writer.write(f)

# --- Entry Point ---
if __name__ == "__main__":
    if GS_EXECUTABLE is None: messagebox.showerror("Critical Error", "Ghostscript not found! Please ensure the executable was built correctly with the Ghostscript directory.")
    else: root = tk.Tk(); app = PdfToolApp(root); root.mainloop()