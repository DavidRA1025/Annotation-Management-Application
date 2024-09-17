import tkinter as tk
from tkinter import ttk
from tkinter import filedialog, messagebox
import pandas as pd
import os

class AnnotationApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Annotation Tool")
        self.annotations = {}
        self.button_column = 0
        self.button_row = 0
        self.current_annotation = None
        self.annotation_file = "annotations.xlsx"

        # Fixed Annotation Display Area with Frame for Buttons (pack only once)
        self.display_frame = ttk.Frame(root)
        self.display_frame.pack(fill=tk.BOTH, expand=True)

        # Button Frame in the display area
        self.display_button_frame = ttk.Frame(self.display_frame)
        self.display_button_frame.pack(side=tk.RIGHT, padx=5, pady=5)

        # Always-On-Top Toggle (within display_button_frame)
        self.always_on_top_var = tk.BooleanVar(value=True) 
        self.always_on_top_checkbox = ttk.Checkbutton(
            self.display_button_frame,
            text="Always On Top",
            variable=self.always_on_top_var,
            command=self.toggle_always_on_top
        )
        self.always_on_top_checkbox.pack() 
        self.root.attributes("-topmost", True)  # Initially set to always on top

        # Display text is packed before the buttons
        self.display_text = tk.Text(self.display_frame, wrap=tk.WORD, height=5, state="disabled")
        self.display_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        buttons = [
            ("Add", self.add_annotation_button),
            ("Delete", self.delete_annotation),
            ("COPY", self.copy_annotation),
            ("Import", self.import_annotations),
            ("Export", self.export_annotations),
            ("Remove All", self.remove_all_annotations)
        ]

        for text, command in buttons:
            ttk.Button(self.display_button_frame, text=text, command=command).pack()

        # Frame for Annotation Buttons
        self.button_frame = ttk.Frame(root)
        self.button_frame.pack(pady=10, fill=tk.BOTH, expand=True)

        self.load_annotations() 

    def toggle_always_on_top(self):  # Corrected function
        new_state = self.always_on_top_var.get()
        self.root.attributes("-topmost", new_state)
   
        # Fixed Annotation Display Area with Frame for Buttons
        self.display_frame = ttk.Frame(root)
        self.display_frame.pack(fill=tk.BOTH, expand=True)

        # Button Frame in the display area
        self.display_button_frame = ttk.Frame(self.display_frame)
        self.display_button_frame.pack(side=tk.RIGHT, padx=5, pady=5)

        # Display text is packed before the buttons
        self.display_text = tk.Text(self.display_frame, wrap=tk.WORD, height=5, state="disabled")
        self.display_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        buttons = [
            ("Add", self.add_annotation_button),
            ("Delete", self.delete_annotation),
            ("COPY", self.copy_annotation),
            ("Import", self.import_annotations),
            ("Export", self.export_annotations),
            ("Remove All", self.remove_all_annotations)
        ]

        for text, command in buttons:
            ttk.Button(self.display_button_frame, text=text, command=command).pack()

        # Frame for Annotation Buttons  (This line was moved to here)
        self.button_frame = ttk.Frame(root)
        self.button_frame.pack(pady=10, fill=tk.BOTH, expand=True)

        self.load_annotations()  # This line was moved to after button_frame initialization

    def add_annotation_button(self):
        top = tk.Toplevel(self.root)
        top.title("Create Annotation")

        ttk.Label(top, text="Name:").grid(row=0, column=0, sticky="w")
        name_entry = ttk.Entry(top)
        name_entry.grid(row=0, column=1, padx=5, pady=5)

        ttk.Label(top, text="Annotation:").grid(row=1, column=0, sticky="w")
        text_entry = tk.Text(top, wrap=tk.WORD, height=3)
        text_entry.grid(row=1, column=1, padx=5, pady=5)

        ttk.Button(top, text="Save", command=lambda: self.save_annotation(top, name_entry.get(), text_entry.get("1.0", tk.END).strip())).grid(row=2, column=0, columnspan=2, pady=10)

    def save_annotation(self, top, name, text): 
        if name and text:
            btn = ttk.Button(self.button_frame, text=name, command=lambda: self.show_annotation(name))
            btn.grid(row=self.button_row, column=self.button_column, sticky="nsew", padx=5, pady=5)
            self.button_frame.columnconfigure(self.button_column, weight=1)
            self.button_frame.rowconfigure(self.button_row, weight=1)
            self.annotations[name] = text
            self.button_column += 1
            if self.button_column > 3:
                self.button_column = 0
                self.button_row += 1
            if top:
                top.destroy() 
            self.save_annotations_to_file()

    def show_annotation(self, name):
        self.current_annotation = name
        self.display_text.config(state="normal")
        self.display_text.delete("1.0", tk.END)
        if name:
            self.display_text.insert(tk.END, self.annotations[name])
        self.display_text.config(state="disabled")

    def copy_annotation(self):
        annotation_text = self.display_text.get("1.0", tk.END).strip()
        if annotation_text:
            self.root.clipboard_clear()
            self.root.clipboard_append(annotation_text)


    def delete_annotation(self):
        if self.current_annotation:
            del self.annotations[self.current_annotation]
            for widget in self.button_frame.winfo_children():
                if isinstance(widget, ttk.Button) and widget["text"] == self.current_annotation:
                    widget.destroy()
                    break
            self.show_annotation(None)
            self.save_annotations_to_file()

    def remove_all_annotations(self):
        confirm = tk.messagebox.askyesno("Confirm Removal", "Are you sure you want to remove all annotations?")
        if confirm:
            self.annotations.clear()
            for widget in self.button_frame.winfo_children():
                widget.destroy()
            self.button_column = 0
            self.button_row = 0
            self.show_annotation(None)

            # Remove the Excel file
            if os.path.exists(self.annotation_file):
                os.remove(self.annotation_file)
                tk.messagebox.showinfo("File Removed", "Annotations and Excel file removed successfully.")
            else:
                tk.messagebox.showinfo("Success", "Annotations removed successfully (file not found).")
    

    def import_annotations(self):  
        filepath = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if filepath:
            try:
                df = pd.read_excel(filepath)
                if 'Name' not in df.columns or 'Annotation' not in df.columns:
                    raise ValueError("Excel file must have 'Name' and 'Annotation' columns")
                for index, row in df.iterrows():
                    name = row['Name']
                    annotation = row['Annotation']
                    if name in self.annotations:
                        overwrite = tk.messagebox.askyesno("Overwrite?", f"Annotation '{name}' already exists. Overwrite?")
                        if not overwrite:
                            continue
                    self.save_annotation(None, name, annotation) 
                self.save_annotations_to_file()
                tk.messagebox.showinfo("Success", "Annotations imported successfully!")
            except (FileNotFoundError, ValueError, pd.errors.EmptyDataError) as e:
                tk.messagebox.showerror("Error", str(e))

    def export_annotations(self):
        filepath = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if filepath:
            try:
                df = pd.DataFrame(list(self.annotations.items()), columns=['Name', 'Annotation'])
                df.to_excel(filepath, index=False)
                tk.messagebox.showinfo("Success", "Annotations exported successfully!")
            except Exception as e:
                tk.messagebox.showerror("Error", f"Failed to export annotations: {e}")

    def load_annotations(self):
        try: 
            if os.path.exists(self.annotation_file):
                df = pd.read_excel(self.annotation_file)
                for index, row in df.iterrows():
                    self.save_annotation(None, row['Name'], row['Annotation'])
        except FileNotFoundError:
            pass 

    def save_annotations_to_file(self):
        if self.annotations:
            df = pd.DataFrame(list(self.annotations.items()), columns=['Name', 'Annotation'])
            df.to_excel(self.annotation_file, index=False)

if __name__ == "__main__":
    root = tk.Tk()
    app = AnnotationApp(root)

    # Center the window
    root.update()
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    window_width = root.winfo_width()
    window_height = root.winfo_height()
    x = (screen_width - window_width) // 2
    y = (screen_height - window_height) // 2
    root.geometry(f"+{x}+{y}")

    # Add Contributor Label
    contributor_label = tk.Label(root, text="contributor: chenwayi@")
    contributor_label.pack(side=tk.BOTTOM, pady=5)

    root.mainloop()