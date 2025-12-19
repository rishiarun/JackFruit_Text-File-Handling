import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import string
from collections import Counter

def normalize_text_for_words(text):
    """Lowercase and remove punctuation (keeps apostrophes optional)."""
    # Use str.translate for fast punctuation removal
    trans_table = str.maketrans(string.punctuation, " " * len(string.punctuation))
    return text.translate(trans_table).lower()

def compute_word_frequency_from_text(text):
    """Return Counter of words from a raw text string."""
    normalized = normalize_text_for_words(text)
    words = [w for w in normalized.split() if w]  # split on whitespace
    return Counter(words)

def load_file_text(filename):
    """Read the whole file and return text; raise exceptions to caller."""
    with open(filename, "r", encoding="utf-8") as f:
        return f.read()

def is_palindrome_core(s):
    """Check palindrome ignoring non-alnum and case."""
    cleaned = "".join(ch.lower() for ch in s if ch.isalnum())
    return cleaned == cleaned[::-1]

def caesar_transform(text, shift):
    """Shift letters by `shift`. Non-letters preserved."""
    result_chars = []
    for ch in text:
        if ch.isalpha():
            base = ord('A') if ch.isupper() else ord('a')
            shifted = chr((ord(ch) - base + shift) % 26 + base)
            result_chars.append(shifted)
        else:
            result_chars.append(ch)
    return "".join(result_chars)

# -------------------------
# GUI callbacks & widgets
# -------------------------
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Text & String Tools")
        self.geometry("700x480")
        self._create_widgets()

    def _create_widgets(self):
        # Notebook with three tabs
        notebook = ttk.Notebook(self)
        notebook.pack(fill="both", expand=True, padx=8, pady=8)

        # --- Word Frequency Tab ---
        wf_frame = ttk.Frame(notebook)
        notebook.add(wf_frame, text="Word Frequency")

        # Filename entry + browse
        file_row = ttk.Frame(wf_frame)
        file_row.pack(fill="x", padx=10, pady=(10, 6))
        ttk.Label(file_row, text="Filename:").pack(side="left")
        self.wf_filename_var = tk.StringVar()
        self.wf_filename_entry = ttk.Entry(file_row, textvariable=self.wf_filename_var)
        self.wf_filename_entry.pack(side="left", fill="x", expand=True, padx=6)
        ttk.Button(file_row, text="Browse...", command=self.wf_browse_file).pack(side="left")

        # Buttons: Load & Analyze, Clear
        btn_row = ttk.Frame(wf_frame)
        btn_row.pack(fill="x", padx=10, pady=(0, 6))
        ttk.Button(btn_row, text="Load & Analyze", command=self.wf_load_and_analyze).pack(side="left")
        ttk.Button(btn_row, text="Clear Output", command=self.wf_clear_output).pack(side="left", padx=8)

        # Output area (Text widget)
        self.wf_output = tk.Text(wf_frame, wrap="word", height=18)
        self.wf_output.pack(fill="both", expand=True, padx=10, pady=(6, 10))

        # --- Palindrome Tab ---
        pal_frame = ttk.Frame(notebook)
        notebook.add(pal_frame, text="Palindrome Checker")

        ttk.Label(pal_frame, text="Enter text to check:").pack(anchor="w", padx=10, pady=(10,4))
        self.pal_input = tk.Entry(pal_frame, width=80)
        self.pal_input.pack(fill="x", padx=10)
        pal_btn_row = ttk.Frame(pal_frame)
        pal_btn_row.pack(fill="x", padx=10, pady=8)
        ttk.Button(pal_btn_row, text="Check", command=self.pal_check).pack(side="left")
        ttk.Button(pal_btn_row, text="Clear", command=lambda: self.pal_input.delete(0, "end")).pack(side="left", padx=6)
        self.pal_result_label = ttk.Label(pal_frame, text="", font=("Segoe UI", 12))
        self.pal_result_label.pack(anchor="w", padx=10, pady=6)

        # --- Caesar Cipher Tab ---
        caesar_frame = ttk.Frame(notebook)
        notebook.add(caesar_frame, text="Caesar Cipher")

        ttk.Label(caesar_frame, text="Input text:").pack(anchor="w", padx=10, pady=(10,4))
        self.c_text = tk.Text(caesar_frame, height=6, wrap="word")
        self.c_text.pack(fill="both", padx=10)

        shift_row = ttk.Frame(caesar_frame)
        shift_row.pack(fill="x", padx=10, pady=8)
        ttk.Label(shift_row, text="Shift (integer):").pack(side="left")
        self.c_shift_var = tk.StringVar(value="3")
        self.c_shift_entry = ttk.Entry(shift_row, textvariable=self.c_shift_var, width=6)
        self.c_shift_entry.pack(side="left", padx=6)

        c_btn_row = ttk.Frame(caesar_frame)
        c_btn_row.pack(fill="x", padx=10)
        ttk.Button(c_btn_row, text="Encrypt →", command=self.c_encrypt).pack(side="left")
        ttk.Button(c_btn_row, text="Decrypt ←", command=self.c_decrypt).pack(side="left", padx=8)
        ttk.Button(c_btn_row, text="Clear", command=lambda: self.c_text.delete("1.0", "end")).pack(side="left", padx=8)

        ttk.Label(caesar_frame, text="Result:").pack(anchor="w", padx=10, pady=(8,0))
        self.c_result = tk.Text(caesar_frame, height=6, wrap="word")
        self.c_result.pack(fill="both", padx=10, pady=(2,10))

        # Menu (optional): Exit and Help
        menubar = tk.Menu(self)
        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="Exit", command=self.quit)
        menubar.add_cascade(label="File", menu=file_menu)
        help_menu = tk.Menu(menubar, tearoff=0)
        help_menu.add_command(label="About", command=self._show_about)
        menubar.add_cascade(label="Help", menu=help_menu)
        self.config(menu=menubar)

    # ------------- Word Frequency callbacks -------------
    def wf_browse_file(self):
        filename = filedialog.askopenfilename(
            title="Select a text file",
            filetypes=[("Text files", "*.txt *.md *.csv *.log"), ("All files", "*.*")]
        )
        if filename:
            self.wf_filename_var.set(filename)

    def wf_load_and_analyze(self):
        filename = self.wf_filename_var.get().strip()
        if not filename:
            messagebox.showwarning("No filename", "Please enter a filename or use Browse...")
            return
        try:
            raw = load_file_text(filename)
        except Exception as e:
            messagebox.showerror("File error", f"Could not open file:\n{e}")
            return

        counter = compute_word_frequency_from_text(raw)
        if not counter:
            self.wf_output.delete("1.0", "end")
            self.wf_output.insert("end", "No words found in file.\n")
            return

        # Format output: most common first
        out_lines = []
        out_lines.append(f"Word frequency for file: {filename}\n")
        out_lines.append(f"{'Word':30}Count")
        out_lines.append("-" * 40)
        for word, cnt in counter.most_common():
            out_lines.append(f"{word:30}{cnt}")
        self.wf_output.delete("1.0", "end")
        self.wf_output.insert("end", "\n".join(out_lines))

    def wf_clear_output(self):
        self.wf_output.delete("1.0", "end")
        self.wf_filename_var.set("")

    # ------------- Palindrome callbacks -------------
    def pal_check(self):
        s = self.pal_input.get()
        if not s.strip():
            self.pal_result_label.config(text="Please enter some text.", foreground="black")
            return
        if is_palindrome_core(s):
            self.pal_result_label.config(text="Palindrome ✓", foreground="green")
        else:
            self.pal_result_label.config(text="Not a palindrome ✗", foreground="red")

    # ------------- Caesar callbacks -------------
    def _get_shift_value(self):
        raw = self.c_shift_var.get().strip()
        try:
            val = int(raw)
            # Normalize shift to range -25..25 for nicer behavior (not required)
            return val % 26
        except ValueError:
            messagebox.showerror("Invalid shift", "Please enter a valid integer for shift.")
            return None

    def c_encrypt(self):
        shift = self._get_shift_value()
        if shift is None:
            return
        text = self.c_text.get("1.0", "end").rstrip("\n")
        if not text:
            messagebox.showinfo("No input", "Enter text to encrypt.")
            return
        result = caesar_transform(text, shift)
        self.c_result.delete("1.0", "end")
        self.c_result.insert("end", result)

    def c_decrypt(self):
        shift = self._get_shift_value()
        if shift is None:
            return
        text = self.c_text.get("1.0", "end").rstrip("\n")
        if not text:
            messagebox.showinfo("No input", "Enter text to decrypt.")
            return
        # decrypt by shifting by negative
        result = caesar_transform(text, -shift)
        self.c_result.delete("1.0", "end")
        self.c_result.insert("end", result)

    def _show_about(self):
        messagebox.showinfo("About", "Text & String Tools\n- Word Frequency\n- Palindrome Checker\n- Caesar Cipher\n\nMade with Tkinter.")

# -------------------------
# Run the app
# -------------------------
if __name__ == "__main__":
    app = App()
    app.mainloop()
