import tkinter as tk
from tkinter import filedialog, messagebox, colorchooser
from tkinter import font as tkfont
import os

# Optional libraries (no errors if missing)
try:
    from docx import Document
except Exception:
    Document = None

try:
    from odf.opendocument import load as odt_load
    from odf.opendocument import OpenDocumentText
    from odf.text import P as OdtP
except Exception:
    odt_load = None
    OpenDocumentText = None
    OdtP = None

SUPPORTED_EXTENSIONS = [
    ("All Supported Files", "*.txt *.md *.doc *.docx *.odt"),
    ("Text Files", "*.txt"),
    ("Markdown Files", "*.md"),
    ("Word Documents", "*.doc *.docx"),
    ("OpenDocument Text", "*.odt"),
]

COMMON_FONTS = [
    "Calibri", "Arial", "Times New Roman", "Courier New", "Verdana", "Tahoma"
]


class TextEditor:
    def __init__(self, root):
        self.root = root
        self.root.title("Multi-Format Text Editor for Windows v0.1")
        self.root.geometry("1100x700")
        self.file_path = None

        self.base_font = tkfont.Font(family="Calibri", size=12)

        # ===== TOP BANNER / TOOLBAR =====
        self.create_banner()

        # ===== TEXT AREA =====
        self.text_area = tk.Text(
            root,
            wrap="word",
            undo=True,
            font=self.base_font
        )
        self.text_area.pack(fill=tk.BOTH, expand=True)

        self.setup_tags()
        self.bind_shortcuts()

    # ================= BANNER =================
    def create_banner(self):
        banner = tk.Frame(self.root, bg="#e6e6e6", height=100, bd=1, relief="raised")
        banner.pack(fill=tk.X, side=tk.TOP)

        # ---- FILE ----
        file_frame = tk.LabelFrame(banner, text="File", padx=8, pady=5)
        file_frame.pack(side=tk.LEFT, padx=8, pady=5)

        tk.Button(file_frame, text="New", width=7, command=self.new_file).pack(side=tk.LEFT, padx=3)
        tk.Button(file_frame, text="Open", width=7, command=self.open_file).pack(side=tk.LEFT, padx=3)
        tk.Button(file_frame, text="Save", width=7, command=self.save_file).pack(side=tk.LEFT, padx=3)

        # ---- FONT FAMILY ----
        font_frame = tk.LabelFrame(banner, text="Font", padx=8, pady=5)
        font_frame.pack(side=tk.LEFT, padx=8, pady=5)

        self.font_family = tk.StringVar(value="Calibri")
        font_dropdown = tk.OptionMenu(font_frame, self.font_family, *COMMON_FONTS, command=self.change_font_family)
        font_dropdown.config(width=12)
        font_dropdown.pack(side=tk.LEFT, padx=3)

        # ---- FONT SIZE ----
        size_frame = tk.LabelFrame(banner, text="Size", padx=8, pady=5)
        size_frame.pack(side=tk.LEFT, padx=8, pady=5)

        self.font_size_var = tk.IntVar(value=12)
        size_box = tk.Spinbox(size_frame, from_=8, to=96, width=5, textvariable=self.font_size_var, command=self.change_font_size)
        size_box.pack(side=tk.LEFT)

        # ---- STYLE (B I U) ----
        style_frame = tk.LabelFrame(banner, text="Style", padx=8, pady=5)
        style_frame.pack(side=tk.LEFT, padx=8, pady=5)

        tk.Button(style_frame, text="B", width=3, font=("Arial", 11, "bold"), command=self.bold).pack(side=tk.LEFT, padx=2)
        tk.Button(style_frame, text="I", width=3, font=("Arial", 11, "italic"), command=self.italic).pack(side=tk.LEFT, padx=2)
        tk.Button(style_frame, text="U", width=3, font=("Arial", 11, "underline"), command=self.underline).pack(side=tk.LEFT, padx=2)

        # ---- ALIGNMENT ----
        align_frame = tk.LabelFrame(banner, text="Alignment", padx=8, pady=5)
        align_frame.pack(side=tk.LEFT, padx=8, pady=5)

        tk.Button(align_frame, text="Left", width=6, command=self.align_left).pack(side=tk.LEFT, padx=2)
        tk.Button(align_frame, text="Center", width=6, command=self.align_center).pack(side=tk.LEFT, padx=2)
        tk.Button(align_frame, text="Right", width=6, command=self.align_right).pack(side=tk.LEFT, padx=2)

        # ---- COLOR PICKER ----
        color_frame = tk.LabelFrame(banner, text="Text Color", padx=8, pady=5)
        color_frame.pack(side=tk.LEFT, padx=8, pady=5)

        tk.Button(color_frame, text="Color", width=8, command=self.choose_color).pack(side=tk.LEFT, padx=2)

    # ================= FORMATTING =================
    def setup_tags(self):
        bold_font = tkfont.Font(self.text_area, self.base_font)
        bold_font.configure(weight="bold")

        italic_font = tkfont.Font(self.text_area, self.base_font)
        italic_font.configure(slant="italic")

        underline_font = tkfont.Font(self.text_area, self.base_font)
        underline_font.configure(underline=1)

        self.text_area.tag_configure("bold", font=bold_font)
        self.text_area.tag_configure("italic", font=italic_font)
        self.text_area.tag_configure("underline", font=underline_font)
        self.text_area.tag_configure("left", justify="left")
        self.text_area.tag_configure("center", justify="center")
        self.text_area.tag_configure("right", justify="right")

    def toggle_tag(self, tag):
        try:
            start = self.text_area.index("sel.first")
            end = self.text_area.index("sel.last")
        except tk.TclError:
            return

        if tag in self.text_area.tag_names("sel.first"):
            self.text_area.tag_remove(tag, start, end)
        else:
            self.text_area.tag_add(tag, start, end)

    def bold(self):
        self.toggle_tag("bold")

    def italic(self):
        self.toggle_tag("italic")

    def underline(self):
        self.toggle_tag("underline")

    def align_left(self):
        self.apply_align("left")

    def align_center(self):
        self.apply_align("center")

    def align_right(self):
        self.apply_align("right")

    def apply_align(self, align_tag):
        try:
            start = self.text_area.index("sel.first")
            end = self.text_area.index("sel.last")
        except tk.TclError:
            start = "1.0"
            end = tk.END

        self.text_area.tag_remove("left", start, end)
        self.text_area.tag_remove("center", start, end)
        self.text_area.tag_remove("right", start, end)
        self.text_area.tag_add(align_tag, start, end)

    def choose_color(self):
        color = colorchooser.askcolor()[1]
        if not color:
            return
        tag_name = f"color_{color}"
        self.text_area.tag_configure(tag_name, foreground=color)

        try:
            start = self.text_area.index("sel.first")
            end = self.text_area.index("sel.last")
            self.text_area.tag_add(tag_name, start, end)
        except tk.TclError:
            pass

    def change_font_family(self, *_):
        self.base_font.configure(family=self.font_family.get())

    def change_font_size(self):
        self.base_font.configure(size=self.font_size_var.get())

    # ================= SHORTCUTS =================
    def bind_shortcuts(self):
        self.root.bind("<Control-n>", lambda e: self.new_file())
        self.root.bind("<Control-o>", lambda e: self.open_file())
        self.root.bind("<Control-s>", lambda e: self.save_file())
        self.root.bind("<Control-b>", lambda e: self.bold())
        self.root.bind("<Control-i>", lambda e: self.italic())
        self.root.bind("<Control-u>", lambda e: self.underline())

    # ================= FILE OPERATIONS =================
    def new_file(self):
        self.text_area.delete(1.0, tk.END)
        self.file_path = None
        self.root.title("Untitled - Python Text Editor")

    def open_file(self):
        path = filedialog.askopenfilename(filetypes=SUPPORTED_EXTENSIONS)
        if not path:
            return
        try:
            content = self.read_file(path)
            self.text_area.delete(1.0, tk.END)
            self.text_area.insert(tk.END, content)
            self.file_path = path
            self.root.title(f"{os.path.basename(path)} - Python Text Editor")
        except Exception:
            messagebox.showerror("Error", "Failed to open file.")

    def save_file(self):
        if not self.file_path:
            self.save_as()
            return
        self.write_file(self.file_path)

    def save_as(self):
        path = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=SUPPORTED_EXTENSIONS
        )
        if not path:
            return
        self.file_path = path
        self.write_file(path)
        self.root.title(f"{os.path.basename(path)} - Python Text Editor")

    # ================= FILE HANDLING =================
    def read_file(self, path):
        ext = os.path.splitext(path)[1].lower()

        if ext in [".txt", ".md", ".doc"]:
            with open(path, "r", encoding="utf-8", errors="ignore") as f:
                return f.read()

        if ext == ".docx" and Document:
            doc = Document(path)
            return "\n".join(p.text for p in doc.paragraphs)

        if ext == ".odt" and odt_load:
            doc = odt_load(path)
            paragraphs = doc.getElementsByType(OdtP)
            return "\n".join(p.firstChild.data if p.firstChild else "" for p in paragraphs)

        with open(path, "r", encoding="utf-8", errors="ignore") as f:
            return f.read()

    def write_file(self, path):
        content = self.text_area.get(1.0, tk.END)
        with open(path, "w", encoding="utf-8") as f:
            f.write(content)


if __name__ == "__main__":
    root = tk.Tk()
    app = TextEditor(root)
    root.mainloop()
