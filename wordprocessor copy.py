import tkinter as tk
from tkinter import filedialog, messagebox, font, colorchooser
from tkinter.simpledialog import askstring
import tkinter.ttk as ttk
from ttkthemes import ThemedTk
from PIL import Image, ImageTk
import os

class WordProcessor:
    def __init__(self, root):
        self.root = root
        self.root.title("ProWrite - Word Processor")
        self.root.geometry("1000x700")
        self.root.minsize(800, 600)

        # Apply modern theme
        self.root.set_theme('arc')

        # Fonts and colors
        self.default_font = ("Segoe UI", 12)
        self.bg_color = "#f5f5f5"
        self.accent_color = "#0078d4"

        # Create main container
        self.main_frame = ttk.Frame(self.root)
        self.main_frame.pack(expand=True, fill='both', padx=10, pady=10)

        # Create toolbar
        self.create_toolbar()

        # Create main frame for text and canvas
        self.content_frame = ttk.Frame(self.main_frame)
        self.content_frame.pack(expand=True, fill='both')

        # Create Text Area
        self.text_area = tk.Text(
            self.content_frame,
            wrap='word',
            font=self.default_font,
            undo=True,
            bd=0,
            relief="flat",
            bg="white",
            fg="black",
            insertbackground="black",
            selectbackground=self.accent_color
        )
        self.text_area.pack(side="left", expand=True, fill='both', padx=(0, 5))

        # Create Canvas for drawing
        self.canvas = tk.Canvas(
            self.content_frame,
            bg="white",
            width=200,
            bd=1,
            relief="flat",
            highlightthickness=1,
            highlightbackground="#d3d3d3"
        )
        self.canvas.pack(side="right", fill="y")

        # Create Menu Bar
        self.create_menu_bar()

        # Status Bar
        self.status_bar = ttk.Label(
            self.root,
            text="Line: 1 | Col: 1 | Words: 0 | Word Wrap: On",
            anchor="w",
            padding=(5, 2)
        )
        self.status_bar.pack(side="bottom", fill="x")

        # Initialize variables
        self.file_path = None
        self.drawing = None
        self.word_wrap = True

        # Bind events
        self.canvas.bind("<ButtonRelease-1>", self.on_mouse_release)
        self.text_area.bind("<KeyRelease>", self.update_status_bar)
        self.text_area.bind("<ButtonRelease-1>", self.update_status_bar)

        # Load icons
        self.icons = self.load_icons()

    def load_icons(self):
        """Load icons for toolbar buttons."""
        icon_dir = "icons"  # Ensure you have an 'icons' folder with PNG images
        icons = {}
        icon_names = ["save", "undo", "cut", "copy", "paste", "bold", "italic", "underline"]
        for name in icon_names:
            try:
                img = Image.open(f"{icon_dir}/{name}.png").resize((20, 20), Image.Resampling.LANCZOS)
                icons[name] = ImageTk.PhotoImage(img)
            except FileNotFoundError:
                icons[name] = None  # Fallback to text if icon not found
        return icons

    def create_toolbar(self):
        """Create a styled toolbar with icons."""
        self.toolbar = ttk.Frame(self.root)
        self.toolbar.pack(side="top", fill="x", padx=5, pady=5)

        buttons = [
            ("Save", self.save_file, "save"),
            ("Undo", self.undo, "undo"),
            ("Cut", self.cut, "cut"),
            ("Copy", self.copy, "copy"),
            ("Paste", self.paste, "paste"),
            ("Bold", self.bold_text, "bold"),
            ("Italic", self.italic_text, "italic"),
            ("Underline", self.underline_text, "underline")
        ]

        for text, command, icon_key in buttons:
            btn = ttk.Button(
                self.toolbar,
                text=text,
                image=self.icons.get(icon_key),
                compound="left",
                command=command
            )
            btn.pack(side="left", padx=2, pady=2)
            btn.bind("<Enter>", lambda e: btn.config(style="Hover.TButton"))
            btn.bind("<Leave>", lambda e: btn.config(style="TButton"))

        # Style for hover effect
        style = ttk.Style()
        style.configure("Hover.TButton", background=self.accent_color)

    def create_menu_bar(self):
        """Create a professional menu bar."""
        self.menu_bar = tk.Menu(self.root)
        self.root.config(menu=self.menu_bar)

        # File menu
        self.file_menu = tk.Menu(self.menu_bar, tearoff=0)
        self.menu_bar.add_cascade(label="File", menu=self.file_menu)
        self.file_menu.add_command(label="New", command=self.new_file, accelerator="Ctrl+N")
        self.file_menu.add_command(label="Open", command=self.open_file, accelerator="Ctrl+O")
        self.file_menu.add_command(label="Save", command=self.save_file, accelerator="Ctrl+S")
        self.file_menu.add_separator()
        self.file_menu.add_command(label="Exit", command=self.root.quit)

        # Edit menu
        self.edit_menu = tk.Menu(self.menu_bar, tearoff=0)
        self.menu_bar.add_cascade(label="Edit", menu=self.edit_menu)
        self.edit_menu.add_command(label="Undo", command=self.undo, accelerator="Ctrl+Z")
        self.edit_menu.add_command(label="Redo", command=self.redo, accelerator="Ctrl+Y")
        self.edit_menu.add_separator()
        self.edit_menu.add_command(label="Cut", command=self.cut, accelerator="Ctrl+X")
        self.edit_menu.add_command(label="Copy", command=self.copy, accelerator="Ctrl+C")
        self.edit_menu.add_command(label="Paste", command=self.paste, accelerator="Ctrl+V")
        self.edit_menu.add_command(label="Find", command=self.search_text, accelerator="Ctrl+F")

        # Format menu
        self.format_menu = tk.Menu(self.menu_bar, tearoff=0)
        self.menu_bar.add_cascade(label="Format", menu=self.format_menu)
        self.format_menu.add_command(label="Bold", command=self.bold_text, accelerator="Ctrl+B")
        self.format_menu.add_command(label="Italic", command=self.italic_text, accelerator="Ctrl+I")
        self.format_menu.add_command(label="Underline", command=self.underline_text, accelerator="Ctrl+U")
        self.format_menu.add_separator()
        self.format_menu.add_command(label="Font", command=self.change_font)
        self.format_menu.add_command(label="Text Color", command=self.change_color)
        self.format_menu.add_separator()
        self.format_menu.add_command(label="Increase Font Size", command=self.increase_font_size)
        self.format_menu.add_command(label="Decrease Font Size", command=self.decrease_font_size)
        self.format_menu.add_command(label="Toggle Word Wrap", command=self.toggle_word_wrap)

        # Draw menu
        self.draw_menu = tk.Menu(self.menu_bar, tearoff=0)
        self.menu_bar.add_cascade(label="Draw", menu=self.draw_menu)
        self.draw_menu.add_command(label="Draw Line", command=self.start_draw_line)
        self.draw_menu.add_command(label="Draw Rectangle", command=self.start_draw_rectangle)
        self.draw_menu.add_command(label="Draw Circle", command=self.start_draw_circle)

        # Bind keyboard shortcuts
        self.root.bind("<Control-n>", lambda e: self.new_file())
        self.root.bind("<Control-o>", lambda e: self.open_file())
        self.root.bind("<Control-s>", lambda e: self.save_file())
        self.root.bind("<Control-z>", lambda e: self.undo())
        self.root.bind("<Control-y>", lambda e: self.redo())
        self.root.bind("<Control-x>", lambda e: self.cut())
        self.root.bind("<Control-c>", lambda e: self.copy())
        self.root.bind("<Control-v>", lambda e: self.paste())
        self.root.bind("<Control-f>", lambda e: self.search_text())
        self.root.bind("<Control-b>", lambda e: self.bold_text())
        self.root.bind("<Control-i>", lambda e: self.italic_text())
        self.root.bind("<Control-u>", lambda e: self.underline_text())

    def new_file(self):
        if self.text_area.get("1.0", "end-1c"):
            if messagebox.askyesno("Save File", "Do you want to save the current document?"):
                self.save_file()
        self.text_area.delete(1.0, "end")
        self.canvas.delete("all")
        self.file_path = None
        self.update_status_bar()

    def open_file(self):
        file_path = filedialog.askopenfilename(
            defaultextension=".txt",
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")]
        )
        if file_path:
            try:
                with open(file_path, "r", encoding="utf-8") as file:
                    content = file.read()
                self.text_area.delete(1.0, "end")
                self.text_area.insert(1.0, content)
                self.file_path = file_path
                self.update_status_bar()
            except Exception as e:
                messagebox.showerror("Error", f"Failed to open file: {e}")

    def save_file(self):
        if not self.file_path:
            self.file_path = filedialog.asksaveasfilename(
                defaultextension=".txt",
                filetypes=[("Text files", "*.txt"), ("All files", "*.*")]
            )
        if self.file_path:
            try:
                with open(self.file_path, "w", encoding="utf-8") as file:
                    content = self.text_area.get("1.0", "end-1c")
                    file.write(content)
                messagebox.showinfo("Success", "File saved successfully")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to save file: {e}")

    def undo(self):
        try:
            self.text_area.edit_undo()
        except Exception:
            pass

    def redo(self):
        try:
            self.text_area.edit_redo()
        except Exception:
            pass

    def cut(self):
        self.text_area.event_generate("<<Cut>>")
        self.update_status_bar()

    def copy(self):
        self.text_area.event_generate("<<Copy>>")

    def paste(self):
        self.text_area.event_generate("<<Paste>>")
        self.update_status_bar()

    def bold_text(self):
        try:
            current_tags = self.text_area.tag_names("sel.first")
            if "bold" in current_tags:
                self.text_area.tag_remove("bold", "sel.first", "sel.last")
            else:
                self.text_area.tag_add("bold", "sel.first", "sel.last")
                bold_font = font.Font(self.text_area, self.text_area.cget("font"))
                bold_font.config(weight="bold")
                self.text_area.tag_configure("bold", font=bold_font)
        except tk.TclError:
            messagebox.showwarning("Warning", "No text selected")

    def italic_text(self):
        try:
            current_tags = self.text_area.tag_names("sel.first")
            if "italic" in current_tags:
                self.text_area.tag_remove("italic", "sel.first", "sel.last")
            else:
                self.text_area.tag_add("italic", "sel.first", "sel.last")
                italic_font = font.Font(self.text_area, self.text_area.cget("font"))
                italic_font.config(slant="italic")
                self.text_area.tag_configure("italic", font=italic_font)
        except tk.TclError:
            messagebox.showwarning("Warning", "No text selected")

    def underline_text(self):
        try:
            current_tags = self.text_area.tag_names("sel.first")
            if "underline" in current_tags:
                self.text_area.tag_remove("underline", "sel.first", "sel.last")
            else:
                self.text_area.tag_add("underline", "sel.first", "sel.last")
                underline_font = font.Font(self.text_area, self.text_area.cget("font"))
                underline_font.config(underline=True)
                self.text_area.tag_configure("underline", font=underline_font)
        except tk.TclError:
            messagebox.showwarning("Warning", "No text selected")

    def change_font(self):
        font_name = askstring("Font", "Enter font name (e.g., Segoe UI):")
        if font_name:
            try:
                font_size = askstring("Font Size", "Enter font size (e.g., 12):")
                font_size = int(font_size)
                self.text_area.config(font=(font_name, font_size))
                self.update_status_bar()
            except ValueError:
                messagebox.showerror("Error", "Invalid font size")

    def change_color(self):
        try:
            color = colorchooser.askcolor(title="Choose Text Color")[1]
            if color:
                self.text_area.tag_add("color", "sel.first", "sel.last")
                self.text_area.tag_configure("color", foreground=color)
        except tk.TclError:
            messagebox.showwarning("Warning", "No text selected")

    def search_text(self):
        search_term = askstring("Search", "Enter search term:")
        if search_term:
            self.text_area.tag_remove("search", "1.0", "end")
            start_pos = "1.0"
            while True:
                start_pos = self.text_area.search(search_term, start_pos, stopindex="end")
                if not start_pos:
                    break
                end_pos = f"{start_pos}+{len(search_term)}c"
                self.text_area.tag_add("search", start_pos, end_pos)
                self.text_area.tag_configure("search", background="#FFFF99")
                start_pos = end_pos

    def increase_font_size(self):
        current_font = font.Font(self.text_area, self.text_area.cget("font"))
        current_size = current_font.actual("size")
        new_size = min(current_size + 2, 72)  # Cap at 72
        self.text_area.config(font=(current_font.actual("family"), new_size))
        self.update_status_bar()

    def decrease_font_size(self):
        current_font = font.Font(self.text_area, self.text_area.cget("font"))
        current_size = current_font.actual("size")
        new_size = max(current_size - 2, 8)  # Minimum size 8
        self.text_area.config(font=(current_font.actual("family"), new_size))
        self.update_status_bar()

    def toggle_word_wrap(self):
        self.word_wrap = not self.word_wrap
        wrap_mode = 'word' if self.word_wrap else 'none'
        self.text_area.config(wrap=wrap_mode)
        self.update_status_bar()

    def start_draw_line(self):
        self.drawing = "line"
        self.canvas.bind("<Button-1>", self.on_mouse_press)

    def start_draw_rectangle(self):
        self.drawing = "rectangle"
        self.canvas.bind("<Button-1>", self.on_mouse_press)

    def start_draw_circle(self):
        self.drawing = "circle"
        self.canvas.bind("<Button-1>", self.on_mouse_press)

    def on_mouse_press(self, event):
        self.start_x = event.x
        self.start_y = event.y

    def on_mouse_release(self, event):
        end_x = event.x
        end_y = event.y
        if self.drawing == "line":
            self.canvas.create_line(self.start_x, self.start_y, end_x, end_y, fill=self.accent_color, width=2)
        elif self.drawing == "rectangle":
            self.canvas.create_rectangle(self.start_x, self.start_y, end_x, end_y, outline=self.accent_color, width=2)
        elif self.drawing == "circle":
            self.canvas.create_oval(self.start_x, self.start_y, end_x, end_y, outline=self.accent_color, width=2)
        self.canvas.unbind("<Button-1>")
        self.drawing = None

    def update_status_bar(self, event=None):
        line, col = self.text_area.index("insert").split(".")
        line = int(line)
        col = int(col)
        words = len(self.text_area.get("1.0", "end-1c").split())
        wrap_status = "On" if self.word_wrap else "Off"
        status = f"Line: {line} | Col: {col} | Words: {words} | Word Wrap: {wrap_status}"
        self.status_bar.config(text=status)

if __name__ == "__main__":
    root = ThemedTk(theme="arc")
    app = WordProcessor(root)
    root.mainloop()