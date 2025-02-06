import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path
import logging
from transcript_generator import TranscriptGenerator

class TranscriptGeneratorGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Lazy Transcript Generator")
        self.root.geometry("600x400")

        self.data_path = tk.StringVar()
        self.writeups_path = tk.StringVar()
        self.output_path = tk.StringVar()
        self.template_path = tk.StringVar()
        self.output_type = tk.StringVar(value='csv')

        self._create_widgets()
        self._setup_layout()

    def _create_widgets(self):
        # File selection frames
        self.input_frame = ttk.LabelFrame(self.root, text="Input Files", padding="10")
        self.output_frame = ttk.LabelFrame(self.root, text="Custom Options", padding="10")

        # Data file selection
        ttk.Label(self.input_frame, text="Data File").grid(row=0, column=0, sticky="w")
        ttk.Entry(self.input_frame, textvariable=self.data_path, width=50).grid(row=0, column=1, padx=5)
        ttk.Button(self.input_frame, text="Browse", command=lambda: self._browse_file(self.data_path, [("CSV files", "*.csv")])).grid(row=0, column=2)

        # Writeups file selection
        ttk.Label(self.input_frame, text="Writeups File:").grid(row=1, column=0, sticky="w")
        ttk.Entry(self.input_frame, textvariable=self.writeups_path, width=50).grid(row=1, column=1, padx=5)
        ttk.Button(self.input_frame, text="Browse", command=lambda: self._browse_file(self.writeups_path, [("CSV files", "*.csv")])).grid(row=1, column=2)

        # Output type selection
        ttk.Label(self.output_frame, text="Output Type:").grid(row=0, column=0, sticky="w")
        ttk.Radiobutton(self.output_frame, text="CSV File", variable=self.output_type, value="csv").grid(row=0, column=1, sticky="w")
        ttk.Radiobutton(self.output_frame, text="DOCX Files", variable=self.output_type, value="docx").grid(row=0, column=2, sticky="w")

        # Template file selection (for DOCX output)
        self.template_label = ttk.Label(self.output_frame, text="Template File:")
        self.template_entry = ttk.Entry(self.output_frame, textvariable=self.template_path, width=50)
        self.template_button = ttk.Button(self.output_frame, text="Browse", command=lambda: self._browse_file(self.template_path, [("Word files", "*.docx")]))

        # Output directory selection
        ttk.Label(self.output_frame, text="Output Location:").grid(row=2, column=0, sticky="w")
        ttk.Entry(self.output_frame, textvariable=self.output_path, width=50).grid(row=2, column=1, columnspan=2, padx=5)
        ttk.Button(self.output_frame, text="Browse", command=self._browse_directory).grid(row=2, column=3)

        # Generate button
        self.generate_button = ttk.Button(self.root, text="Generate Transcripts", command=self.generate_transcripts)

        self.output_type.trace_add('write', self._toggle_template_widgets)

    def _setup_layout(self):
        self.root.grid_columnconfigure(0, weight=1)
        self.root.grid_rowconfigure(2, weight=1)

        self.input_frame.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")
        self.output_frame.grid(row=1, column=0, padx=10, pady=10, sticky="nsew")
        self.generate_button.grid(row=2, column=0, pady=20)

        self.input_frame.grid_columnconfigure(1, weight=1)
        self.output_frame.grid_columnconfigure(1, weight=1)

    def _toggle_template_widgets(self, var_name, index, mode):
        if self.output_type.get() == 'docx':
            self.template_label.grid(row=1, column=0, sticky="w")
            self.template_entry.grid(row=1, column=1, columnspan=2, padx=5)
            self.template_button.grid(row=1, column=3)
        else:
            self.template_label.grid_remove()
            self.template_entry.grid_remove()
            self.template_button.grid_remove()
    
    def _browse_file(self, string_var, file_types):
        filename = filedialog.askopenfilename(filetypes=file_types)
        if filename:
            string_var.set(filename)
        
    def _browse_directory(self):
        directory = filedialog.askdirectory()
        if directory:
            self.output_path.set(directory)

    def generate_transcripts(self):
        try:
            if not self.data_path.get() or not self.writeups_path.get():
                messagebox.showerror("Error", "Please select both data and writeups file")
            if self.output_type.get() == "docx" and not self.template_path.get():
                messagebox.showerror("Error", "Please select a template file for DOCX output")
            
            generator = TranscriptGenerator(self.data_path.get(), self.writeups_path.get())

            if self.output_type.get() == 'csv':
                output_path = generator.export_transcripts_csv(self.output_path.get())
                messagebox.showinfo("Success", f"CSV file generated successfully at:\n{output_path}")
            
            else: 
                output_dir = generator.export_from_docx_template(self.template_path.get(), self.output_path.get())
                messagebox.showinfo("Success", f"DOCX files generated successfully in:\n{output_dir}")

        except Exception as e:
            messagebox.showerror("Error", f"An error occurred:\n{str(e)}")
            logging.error(f"Error generating transcripts: {e}")

def main():
    logging.basicConfig(level=logging.INFO)
    root = tk.Tk()
    app = TranscriptGeneratorGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()