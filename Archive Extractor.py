import tkinter as tk
from tkinter import filedialog, messagebox
import subprocess
import os
import zipfile
import py7zr

class ArchiveExtractor:
    def __init__(self, master):
        self.master = master
        self.master.title("Multi Archive Password Extractor")
        
        self.file_paths = []
        self.output_dir = os.getcwd()  # Default directory is the current folder
        
        # Frame for file selection
        self.frame = tk.Frame(self.master)
        self.frame.pack(pady=20)
        
        self.label = tk.Label(self.frame, text="Select files (RAR, ZIP, 7z):")
        self.label.pack(side=tk.LEFT)
        
        self.file_entry = tk.Entry(self.frame, width=50)
        self.file_entry.pack(side=tk.LEFT, padx=5)
        
        self.browse_button = tk.Button(self.frame, text="Browse", command=self.browse_files)
        self.browse_button.pack(side=tk.LEFT)

        # Frame for output folder selection
        self.output_frame = tk.Frame(self.master)
        self.output_frame.pack(pady=10)

        self.output_label = tk.Label(self.output_frame, text="Choose extraction path:")
        self.output_label.pack(side=tk.LEFT)
        
        self.output_entry = tk.Entry(self.output_frame, width=50)
        self.output_entry.insert(0, self.output_dir)  # Set default path
        self.output_entry.pack(side=tk.LEFT, padx=5)
        
        self.output_browse_button = tk.Button(self.output_frame, text="Browse", command=self.browse_output)
        self.output_browse_button.pack(side=tk.LEFT)

        # Input for passwords
        self.password_label = tk.Label(self.master, text="Enter Passwords (one per line):")
        self.password_label.pack(pady=10)

        self.password_text = tk.Text(self.master, width=60, height=10)
        self.password_text.pack(pady=5)

        # Button to extract
        self.extract_button = tk.Button(self.master, text="Extract", command=self.extract)
        self.extract_button.pack(pady=20)

    def browse_files(self):
        self.file_paths = filedialog.askopenfilenames(filetypes=[("Archive files", "*.rar;*.zip;*.7z")])
        self.file_entry.delete(0, tk.END)
        self.file_entry.insert(0, ', '.join(self.file_paths))

    def browse_output(self):
        # Opens a dialog to select the output folder
        folder_selected = filedialog.askdirectory(initialdir=self.output_dir)
        if folder_selected:
            self.output_dir = folder_selected
            self.output_entry.delete(0, tk.END)
            self.output_entry.insert(0, self.output_dir)

    def extract(self):
        if not self.file_paths:
            messagebox.showerror("Error", "You must select at least one file.")
            return

        # Get passwords from the Text widget
        passwords = self.password_text.get("1.0", tk.END).strip().splitlines()
        passwords = [p.strip() for p in passwords if p.strip()]  # Remove empty spaces

        success_count = 0
        failure_count = 0

        for file_path in self.file_paths:
            extracted = False
            try:
                extension = os.path.splitext(file_path)[-1].lower()
                if extension == '.rar':
                    extracted = self.extract_rar(file_path, passwords, self.output_dir)
                elif extension == '.zip':
                    extracted = self.extract_zip(file_path, passwords, self.output_dir)
                elif extension == '.7z':
                    extracted = self.extract_7z(file_path, passwords, self.output_dir)
                else:
                    messagebox.showerror("Error", f"Unsupported file format: {file_path}")

                if extracted:
                    success_count += 1
                else:
                    failure_count += 1

            except Exception as e:
                messagebox.showerror("Error", str(e))

        # Show final result
        messagebox.showinfo("Result", f"Successfully extracted: {success_count}\nFailed to extract: {failure_count}")

    def extract_rar(self, file_path, passwords, output_dir):
        # Try extracting without a password first, if no passwords are specified
        if not passwords:
            passwords = ['']  # Add an empty password to attempt extraction without a password

        for password in passwords:
            try:
                # Command to extract the rar file with WinRAR
                extract_command = [r'C:\Program Files\WinRAR\WinRAR.exe', 'x', '-p' + password, '-y', file_path, output_dir]
                result = subprocess.run(extract_command, capture_output=True, text=True)
                
                if result.returncode == 0:  # If the extraction was successful
                    return True
                else:
                    # Check if the error relates to a wrong password
                    if "password" not in result.stderr.lower() and "wrong" not in result.stderr.lower():
                        print(result.stderr)  # Print the error for debugging
            except Exception as e:
                continue  # Ignore exceptions and try the next password
        return False  # Failed

    def extract_zip(self, file_path, passwords, output_dir):
        # Try extracting without a password first
        if not passwords:
            passwords = ['']  # Add an empty password to attempt extraction without a password

        for password in passwords:
            try:
                with zipfile.ZipFile(file_path) as zf:
                    zf.extractall(pwd=password.encode('utf-8'), path=output_dir)
                return True  # Successfully extracted
            except (RuntimeError, zipfile.BadZipFile):
                continue  
        return False  # Failed

    def extract_7z(self, file_path, passwords, output_dir):
        # Try extracting without a password first
        if not passwords:
            passwords = ['']  # Add an empty password to attempt extraction without a password

        for password in passwords:
            try:
                with py7zr.SevenZipFile(file_path, mode='r', password=password) as sz:
                    sz.extractall(path=output_dir)
                return True  # Successfully extracted
            except Exception as e:
                continue 
        return False  # Failed

if __name__ == "__main__":
    root = tk.Tk()
    extractor = ArchiveExtractor(root)
    root.mainloop()