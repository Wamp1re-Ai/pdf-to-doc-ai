#!/usr/bin/env python3
"""
PDF to Word Converter GUI
Simple graphical interface for converting PDF files to Word documents using Gemini API
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import threading
import os
from pathlib import Path
import sys

# Import our converter class
try:
    from pdf_to_word_converter import PDFToWordConverter
except ImportError:
    messagebox.showerror("Error", "pdf_to_word_converter.py not found in the same directory!")
    sys.exit(1)

class PDFToWordGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF to Word Converter")
        self.root.geometry("600x500")
        self.root.resizable(True, True)
        
        # Variables
        self.pdf_file_path = tk.StringVar()
        self.output_file_path = tk.StringVar()
        self.api_key = tk.StringVar()
        self.is_converting = False
        
        # Load API key from environment if available
        env_api_key = os.getenv('GEMINI_API_KEY', '')
        self.api_key.set(env_api_key)
        
        self.setup_ui()
        
    def setup_ui(self):
        """Setup the user interface"""
        # Main frame
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        
        # Title
        title_label = ttk.Label(main_frame, text="PDF to Word Converter", 
                               font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        # API Key section
        ttk.Label(main_frame, text="Gemini API Key:").grid(row=1, column=0, sticky=tk.W, pady=5)
        api_key_entry = ttk.Entry(main_frame, textvariable=self.api_key, show="*", width=50)
        api_key_entry.grid(row=1, column=1, sticky=(tk.W, tk.E), padx=(10, 5), pady=5)
        
        ttk.Button(main_frame, text="Help", command=self.show_api_help, width=8).grid(
            row=1, column=2, padx=(5, 0), pady=5)
        
        # Input file section
        ttk.Label(main_frame, text="PDF File:").grid(row=2, column=0, sticky=tk.W, pady=5)
        pdf_entry = ttk.Entry(main_frame, textvariable=self.pdf_file_path, width=50)
        pdf_entry.grid(row=2, column=1, sticky=(tk.W, tk.E), padx=(10, 5), pady=5)
        
        ttk.Button(main_frame, text="Browse", command=self.browse_pdf_file, width=8).grid(
            row=2, column=2, padx=(5, 0), pady=5)
        
        # Output file section
        ttk.Label(main_frame, text="Output File:").grid(row=3, column=0, sticky=tk.W, pady=5)
        output_entry = ttk.Entry(main_frame, textvariable=self.output_file_path, width=50)
        output_entry.grid(row=3, column=1, sticky=(tk.W, tk.E), padx=(10, 5), pady=5)
        
        ttk.Button(main_frame, text="Browse", command=self.browse_output_file, width=8).grid(
            row=3, column=2, padx=(5, 0), pady=5)
        
        # Convert button
        self.convert_button = ttk.Button(main_frame, text="Convert PDF to Word", 
                                        command=self.start_conversion, style="Accent.TButton")
        self.convert_button.grid(row=4, column=0, columnspan=3, pady=20)
        
        # Progress bar
        self.progress = ttk.Progressbar(main_frame, mode='indeterminate')
        self.progress.grid(row=5, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)
        
        # Status label
        self.status_label = ttk.Label(main_frame, text="Ready to convert", foreground="green")
        self.status_label.grid(row=6, column=0, columnspan=3, pady=5)
        
        # Log area
        ttk.Label(main_frame, text="Conversion Log:").grid(row=7, column=0, sticky=tk.W, pady=(20, 5))
        
        self.log_text = scrolledtext.ScrolledText(main_frame, height=12, width=70)
        self.log_text.grid(row=8, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        
        # Configure grid weights for resizing
        main_frame.rowconfigure(8, weight=1)
        
        # Add some initial text to log
        self.log_message("Welcome to PDF to Word Converter!")
        self.log_message("Please enter your Gemini API key and select a PDF file to convert.")
        
    def browse_pdf_file(self):
        """Open file dialog to select PDF file"""
        file_path = filedialog.askopenfilename(
            title="Select PDF File",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
        )
        if file_path:
            self.pdf_file_path.set(file_path)
            # Auto-generate output filename
            pdf_name = Path(file_path).stem
            output_path = str(Path(file_path).parent / f"{pdf_name}_converted.docx")
            self.output_file_path.set(output_path)
            self.log_message(f"Selected PDF: {file_path}")
    
    def browse_output_file(self):
        """Open file dialog to select output location"""
        file_path = filedialog.asksaveasfilename(
            title="Save Word Document As",
            defaultextension=".docx",
            filetypes=[("Word documents", "*.docx"), ("All files", "*.*")]
        )
        if file_path:
            self.output_file_path.set(file_path)
            self.log_message(f"Output will be saved to: {file_path}")
    
    def show_api_help(self):
        """Show help dialog for API key"""
        help_text = """To get a Gemini API key:

1. Go to Google AI Studio: https://makersuite.google.com/app/apikey
2. Sign in with your Google account
3. Click "Create API Key"
4. Copy the generated key and paste it here

You can also set the GEMINI_API_KEY environment variable to avoid entering it each time."""
        
        messagebox.showinfo("API Key Help", help_text)
    
    def log_message(self, message):
        """Add message to log area"""
        self.log_text.insert(tk.END, f"{message}\n")
        self.log_text.see(tk.END)
        self.root.update_idletasks()
    
    def update_status(self, message, color="black"):
        """Update status label"""
        self.status_label.config(text=message, foreground=color)
        self.root.update_idletasks()
    
    def validate_inputs(self):
        """Validate user inputs before conversion"""
        if not self.api_key.get().strip():
            messagebox.showerror("Error", "Please enter your Gemini API key")
            return False
        
        if not self.pdf_file_path.get().strip():
            messagebox.showerror("Error", "Please select a PDF file")
            return False
        
        if not os.path.exists(self.pdf_file_path.get()):
            messagebox.showerror("Error", "Selected PDF file does not exist")
            return False
        
        if not self.output_file_path.get().strip():
            messagebox.showerror("Error", "Please specify an output file location")
            return False
        
        return True
    
    def start_conversion(self):
        """Start the conversion process in a separate thread"""
        if self.is_converting:
            return
        
        if not self.validate_inputs():
            return
        
        self.is_converting = True
        self.convert_button.config(state="disabled", text="Converting...")
        self.progress.start()
        self.update_status("Converting...", "blue")
        
        # Start conversion in separate thread to prevent GUI freezing
        thread = threading.Thread(target=self.perform_conversion)
        thread.daemon = True
        thread.start()
    
    def perform_conversion(self):
        """Perform the actual conversion"""
        try:
            self.log_message("Starting conversion process...")

            # Initialize converter
            converter = PDFToWordConverter(self.api_key.get().strip())
            self.log_message("Gemini API initialized successfully")
            self.log_message(f"Using model: {converter.model_names[0]} (with fallback support)")
            
            # Perform conversion
            output_file = converter.convert_pdf_to_word(
                pdf_path=self.pdf_file_path.get(),
                output_path=self.output_file_path.get()
            )
            
            # Success
            self.root.after(0, self.conversion_success, output_file)
            
        except Exception as e:
            # Error
            self.root.after(0, self.conversion_error, str(e))
    
    def conversion_success(self, output_file):
        """Handle successful conversion"""
        self.is_converting = False
        self.convert_button.config(state="normal", text="Convert PDF to Word")
        self.progress.stop()
        self.update_status("Conversion completed successfully!", "green")
        
        self.log_message("✅ Conversion completed successfully!")
        self.log_message(f"📝 Word document saved to: {output_file}")
        
        # Ask if user wants to open the file
        if messagebox.askyesno("Success", f"Conversion completed!\n\nOpen the Word document now?"):
            try:
                os.startfile(output_file)  # Windows
            except AttributeError:
                try:
                    os.system(f'open "{output_file}"')  # macOS
                except:
                    os.system(f'xdg-open "{output_file}"')  # Linux
    
    def conversion_error(self, error_message):
        """Handle conversion error"""
        self.is_converting = False
        self.convert_button.config(state="normal", text="Convert PDF to Word")
        self.progress.stop()
        self.update_status("Conversion failed", "red")
        
        self.log_message(f"❌ Conversion failed: {error_message}")
        messagebox.showerror("Conversion Error", f"Conversion failed:\n\n{error_message}")

def main():
    """Main function to run the GUI"""
    root = tk.Tk()
    
    # Set up styling
    style = ttk.Style()
    style.theme_use('clam')  # Use a modern theme
    
    app = PDFToWordGUI(root)
    
    # Center the window
    root.update_idletasks()
    x = (root.winfo_screenwidth() // 2) - (root.winfo_width() // 2)
    y = (root.winfo_screenheight() // 2) - (root.winfo_height() // 2)
    root.geometry(f"+{x}+{y}")
    
    root.mainloop()

if __name__ == "__main__":
    main()
