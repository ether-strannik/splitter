"""
Panorama Slicer - Cut large panoramas into printable 8.5x11" landscape pages
Navigation: Mouse drag to pan, scroll wheel to zoom, arrow keys to pan
Click a page cell to export it, or use buttons to export all/visible pages
"""

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from PIL import Image, ImageTk, ImageDraw, ImageWin
import os
import math

try:
    import win32print
    import win32ui
    import win32con
    HAS_WIN32 = True
except ImportError:
    HAS_WIN32 = False

# Page dimensions in inches (landscape)
PAGE_WIDTH_INCHES = 11
PAGE_HEIGHT_INCHES = 8.5


class PanoramaSlicer:
    def __init__(self, root):
        self.root = root
        self.root.title("Panorama Slicer")
        self.root.geometry("1400x900")

        # Image state
        self.original_image = None
        self.image_path = None
        self.img_width = 0
        self.img_height = 0

        # View state
        self.zoom = 0.1  # Start zoomed out for large images
        self.pan_x = 0  # Pan offset in image coordinates
        self.pan_y = 0
        self.drag_start_x = 0
        self.drag_start_y = 0
        self.is_dragging = False

        # Grid settings
        self.show_grid = tk.BooleanVar(value=True)
        self.right_to_left = tk.BooleanVar(value=True)  # Start from right end
        self.start_offset = 0  # Custom start position (from right if RTL, from left if LTR)
        self.grid_color = "#FF0000"
        self.grid_alpha = 128

        # Output settings
        self.output_dir = None

        self.setup_ui()
        self.bind_events()

    def setup_ui(self):
        # Top toolbar
        toolbar = ttk.Frame(self.root)
        toolbar.pack(fill=tk.X, padx=5, pady=5)

        ttk.Button(toolbar, text="Open Image", command=self.open_image).pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="Set Output Dir", command=self.set_output_dir).pack(side=tk.LEFT, padx=2)

        ttk.Separator(toolbar, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=10)

        ttk.Checkbutton(toolbar, text="Show Grid", variable=self.show_grid,
                        command=self.refresh_view).pack(side=tk.LEFT, padx=2)
        ttk.Checkbutton(toolbar, text="Right-to-Left", variable=self.right_to_left,
                        command=self.on_direction_change).pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="Reset Start", command=self.reset_start).pack(side=tk.LEFT, padx=2)

        ttk.Separator(toolbar, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=10)

        ttk.Button(toolbar, text="Export All Pages", command=self.export_all_pages).pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="Export Visible", command=self.export_visible_pages).pack(side=tk.LEFT, padx=2)

        ttk.Separator(toolbar, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=10)

        ttk.Button(toolbar, text="Print...", command=self.show_print_dialog).pack(side=tk.LEFT, padx=2)

        ttk.Separator(toolbar, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=10)

        ttk.Label(toolbar, text="Zoom:").pack(side=tk.LEFT, padx=2)
        self.zoom_label = ttk.Label(toolbar, text="10%")
        self.zoom_label.pack(side=tk.LEFT, padx=2)

        ttk.Button(toolbar, text="Fit", command=self.fit_to_window).pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="100%", command=self.zoom_100).pack(side=tk.LEFT, padx=2)

        # Info panel
        info_frame = ttk.Frame(self.root)
        info_frame.pack(fill=tk.X, padx=5)

        self.info_label = ttk.Label(info_frame, text="No image loaded")
        self.info_label.pack(side=tk.LEFT)

        self.page_info_label = ttk.Label(info_frame, text="")
        self.page_info_label.pack(side=tk.RIGHT)

        # Canvas
        canvas_frame = ttk.Frame(self.root)
        canvas_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        self.canvas = tk.Canvas(canvas_frame, bg="#2a2a2a", highlightthickness=0)
        self.canvas.pack(fill=tk.BOTH, expand=True)

        # Status bar
        self.status_label = ttk.Label(self.root, text="Ready - Open an image to begin")
        self.status_label.pack(fill=tk.X, padx=5, pady=2)

    def bind_events(self):
        self.canvas.bind("<ButtonPress-1>", self.on_mouse_down)
        self.canvas.bind("<B1-Motion>", self.on_mouse_drag)
        self.canvas.bind("<ButtonRelease-1>", self.on_mouse_up)
        self.canvas.bind("<ButtonPress-3>", self.on_right_click)  # Right-click to set start
        self.canvas.bind("<MouseWheel>", self.on_scroll)  # Windows
        self.canvas.bind("<Button-4>", self.on_scroll)  # Linux scroll up
        self.canvas.bind("<Button-5>", self.on_scroll)  # Linux scroll down
        self.canvas.bind("<Configure>", self.on_resize)
        self.canvas.bind("<Motion>", self.on_mouse_move)

        # Arrow keys
        self.root.bind("<Left>", lambda e: self.pan_by(-100, 0))
        self.root.bind("<Right>", lambda e: self.pan_by(100, 0))
        self.root.bind("<Up>", lambda e: self.pan_by(0, -100))
        self.root.bind("<Down>", lambda e: self.pan_by(0, 100))

        # Shift+arrows for faster pan
        self.root.bind("<Shift-Left>", lambda e: self.pan_by(-500, 0))
        self.root.bind("<Shift-Right>", lambda e: self.pan_by(500, 0))
        self.root.bind("<Shift-Up>", lambda e: self.pan_by(0, -500))
        self.root.bind("<Shift-Down>", lambda e: self.pan_by(0, 500))

    def open_image(self):
        path = filedialog.askopenfilename(
            filetypes=[("Image files", "*.png *.jpg *.jpeg *.tif *.tiff *.bmp"), ("All files", "*.*")]
        )
        if path:
            self.load_image(path)

    def load_image(self, path):
        self.status_label.config(text=f"Loading {os.path.basename(path)}...")
        self.root.update()

        try:
            # Load full image
            self.original_image = Image.open(path)
            self.original_image.load()  # Force load into memory
            self.image_path = path
            self.img_width, self.img_height = self.original_image.size

            # Calculate page dimensions based on image height = 8.5 inches
            # So pixels_per_inch = img_height / 8.5, and page_width = ppi * 11
            self.pixels_per_inch = self.img_height / PAGE_HEIGHT_INCHES
            self.page_width_px = int(self.pixels_per_inch * PAGE_WIDTH_INCHES)
            self.page_height_px = self.img_height  # Full height fits on page

            # Reset start offset
            self.start_offset = 0

            # Calculate page grid info (only horizontal slicing needed)
            self.pages_x = math.ceil(self.img_width / self.page_width_px)
            self.pages_y = 1  # Height already fits
            total_pages = self.pages_x

            # Update info
            self.info_label.config(
                text=f"{os.path.basename(path)} | {self.img_width}x{self.img_height}px | "
                     f"PPI: {self.pixels_per_inch:.0f} | Page: {self.page_width_px}x{self.page_height_px}px | "
                     f"{total_pages} pages"
            )

            # Set default output dir
            if not self.output_dir:
                self.output_dir = os.path.dirname(path)

            # Fit to window
            self.fit_to_window()

            self.status_label.config(text=f"Loaded. Click on grid cell to export single page, or use Export buttons.")

        except Exception as e:
            messagebox.showerror("Error", f"Failed to load image: {e}")
            self.status_label.config(text="Error loading image")

    def fit_to_window(self):
        if not self.original_image:
            return
        canvas_w = self.canvas.winfo_width()
        canvas_h = self.canvas.winfo_height()
        if canvas_w < 10 or canvas_h < 10:
            canvas_w, canvas_h = 1400, 800

        zoom_x = canvas_w / self.img_width
        zoom_y = canvas_h / self.img_height
        self.zoom = min(zoom_x, zoom_y) * 0.95
        self.pan_x = 0
        self.pan_y = 0
        self.refresh_view()

    def zoom_100(self):
        if not self.original_image:
            return
        self.zoom = 1.0
        self.refresh_view()

    def set_output_dir(self):
        dir_path = filedialog.askdirectory(initialdir=self.output_dir)
        if dir_path:
            self.output_dir = dir_path
            self.status_label.config(text=f"Output directory: {dir_path}")

    def refresh_view(self):
        if not self.original_image:
            return

        self.canvas.delete("all")

        canvas_w = self.canvas.winfo_width()
        canvas_h = self.canvas.winfo_height()

        # Calculate visible region in image coordinates
        view_w = canvas_w / self.zoom
        view_h = canvas_h / self.zoom

        # Clamp pan
        max_pan_x = max(0, self.img_width - view_w)
        max_pan_y = max(0, self.img_height - view_h)
        self.pan_x = max(0, min(self.pan_x, max_pan_x))
        self.pan_y = max(0, min(self.pan_y, max_pan_y))

        # Crop region
        left = int(self.pan_x)
        top = int(self.pan_y)
        right = int(min(self.pan_x + view_w, self.img_width))
        bottom = int(min(self.pan_y + view_h, self.img_height))

        # Extract and resize visible portion
        crop = self.original_image.crop((left, top, right, bottom))
        display_w = int((right - left) * self.zoom)
        display_h = int((bottom - top) * self.zoom)

        if display_w > 0 and display_h > 0:
            resized = crop.resize((display_w, display_h), Image.Resampling.NEAREST if self.zoom > 0.5 else Image.Resampling.BILINEAR)

            # Draw grid overlay if enabled
            if self.show_grid.get():
                resized = resized.convert("RGBA")
                overlay = Image.new("RGBA", resized.size, (0, 0, 0, 0))
                draw = ImageDraw.Draw(overlay)

                # Calculate grid lines in view coordinates (vertical lines for page boundaries)
                if self.right_to_left.get():
                    # RTL: grid lines start from custom start point (or right edge)
                    start_x = self.img_width - self.start_offset
                    for px in range(self.pages_x + 1):
                        x_img = start_x - px * self.page_width_px
                        x_view = (x_img - self.pan_x) * self.zoom
                        if 0 <= x_view <= display_w:
                            # First line (start point) in green, others in red
                            color = (0, 255, 0, 200) if px == 0 else (255, 0, 0, 200)
                            draw.line([(x_view, 0), (x_view, display_h)], fill=color, width=2)
                else:
                    # LTR: grid lines start from custom start point (or left edge)
                    start_x = self.start_offset
                    for px in range(self.pages_x + 1):
                        x_img = start_x + px * self.page_width_px
                        x_view = (x_img - self.pan_x) * self.zoom
                        if 0 <= x_view <= display_w:
                            color = (0, 255, 0, 200) if px == 0 else (255, 0, 0, 200)
                            draw.line([(x_view, 0), (x_view, display_h)], fill=color, width=2)

                # Draw page numbers
                for px in range(self.pages_x):
                    page_num = px + 1
                    if self.right_to_left.get():
                        start_x = self.img_width - self.start_offset
                        x_img = start_x - px * self.page_width_px - self.page_width_px // 2
                    else:
                        start_x = self.start_offset
                        x_img = start_x + px * self.page_width_px + self.page_width_px // 2

                    y_img = self.page_height_px // 2
                    x_view = (x_img - self.pan_x) * self.zoom
                    y_view = (y_img - self.pan_y) * self.zoom
                    if 0 <= x_view <= display_w and 0 <= y_view <= display_h:
                        text = f"{page_num}"
                        draw.rectangle([x_view-20, y_view-10, x_view+20, y_view+10],
                                     fill=(0, 0, 0, 150))
                        draw.text((x_view-10, y_view-8), text, fill=(255, 255, 255, 255))

                resized = Image.alpha_composite(resized, overlay)
                resized = resized.convert("RGB")

            self.display_image = ImageTk.PhotoImage(resized)
            self.canvas.create_image(0, 0, anchor=tk.NW, image=self.display_image)

        self.zoom_label.config(text=f"{self.zoom*100:.1f}%")

    def on_mouse_down(self, event):
        self.drag_start_x = event.x
        self.drag_start_y = event.y
        self.drag_pan_x = self.pan_x
        self.drag_pan_y = self.pan_y
        self.is_dragging = False

    def on_mouse_drag(self, event):
        dx = event.x - self.drag_start_x
        dy = event.y - self.drag_start_y

        if abs(dx) > 5 or abs(dy) > 5:
            self.is_dragging = True

        self.pan_x = self.drag_pan_x - dx / self.zoom
        self.pan_y = self.drag_pan_y - dy / self.zoom
        self.refresh_view()

    def on_mouse_up(self, event):
        if not self.is_dragging and self.original_image:
            # Click - export the clicked page
            self.export_clicked_page(event.x, event.y)

    def on_scroll(self, event):
        if not self.original_image:
            return

        # Get scroll direction
        if event.num == 4 or (hasattr(event, 'delta') and event.delta > 0):
            factor = 1.2
        else:
            factor = 1 / 1.2

        # Zoom centered on mouse position
        mouse_x_img = self.pan_x + event.x / self.zoom
        mouse_y_img = self.pan_y + event.y / self.zoom

        old_zoom = self.zoom
        self.zoom = max(0.01, min(5.0, self.zoom * factor))

        # Adjust pan to keep mouse position stable
        self.pan_x = mouse_x_img - event.x / self.zoom
        self.pan_y = mouse_y_img - event.y / self.zoom

        self.refresh_view()

    def on_resize(self, event):
        self.refresh_view()

    def on_mouse_move(self, event):
        if not self.original_image:
            return

        # Calculate which page the mouse is over
        img_x = self.pan_x + event.x / self.zoom

        if self.right_to_left.get():
            start_x = self.img_width - self.start_offset
            dist_from_start = start_x - img_x
        else:
            start_x = self.start_offset
            dist_from_start = img_x - start_x

        page_num = int(dist_from_start // self.page_width_px) + 1

        if 1 <= page_num <= self.pages_x:
            self.page_info_label.config(text=f"Page {page_num} of {self.pages_x} | L-click: export | R-click: set start")
        else:
            self.page_info_label.config(text="R-click: set start here")

    def on_right_click(self, event):
        """Right-click to set where page 1 starts."""
        if not self.original_image:
            return

        img_x = self.pan_x + event.x / self.zoom

        if self.right_to_left.get():
            # RTL: set the right edge of page 1
            self.start_offset = self.img_width - img_x
        else:
            # LTR: set the left edge of page 1
            self.start_offset = img_x

        # Clamp to valid range
        self.start_offset = max(0, min(self.start_offset, self.img_width - self.page_width_px))

        # Recalculate number of pages
        remaining_width = self.img_width - self.start_offset
        self.pages_x = math.ceil(remaining_width / self.page_width_px)

        self.status_label.config(text=f"Start point set. {self.pages_x} pages from this position.")
        self.refresh_view()

    def reset_start(self):
        """Reset start point to edge."""
        self.start_offset = 0
        if self.original_image:
            self.pages_x = math.ceil(self.img_width / self.page_width_px)
        self.refresh_view()
        self.status_label.config(text="Start point reset to edge.")

    def on_direction_change(self):
        """Reset start when direction changes."""
        self.reset_start()

    def pan_by(self, dx, dy):
        if not self.original_image:
            return
        self.pan_x += dx / self.zoom
        self.pan_y += dy / self.zoom
        self.refresh_view()

    def export_clicked_page(self, canvas_x, canvas_y):
        if not self.original_image or not self.output_dir:
            messagebox.showwarning("Warning", "Please load an image and set output directory first")
            return

        img_x = self.pan_x + canvas_x / self.zoom

        if self.right_to_left.get():
            start_x = self.img_width - self.start_offset
            dist_from_start = start_x - img_x
        else:
            start_x = self.start_offset
            dist_from_start = img_x - start_x

        page_num = int(dist_from_start // self.page_width_px) + 1

        if 1 <= page_num <= self.pages_x:
            self.export_page(page_num)

    def export_page(self, page_num):
        """Export a page by page number (1-based)."""

        if self.right_to_left.get():
            # RTL: page 1 starts at start_offset from right edge
            start_x = self.img_width - self.start_offset
            right = start_x - (page_num - 1) * self.page_width_px
            left = max(0, right - self.page_width_px)
        else:
            # LTR: page 1 starts at start_offset from left edge
            start_x = self.start_offset
            left = start_x + (page_num - 1) * self.page_width_px
            right = min(left + self.page_width_px, self.img_width)

        top = 0
        bottom = self.img_height

        # Crop the page
        page_img = self.original_image.crop((left, top, right, bottom))

        # If edge page is narrower, pad with white to maintain exact page width
        if page_img.size[0] != self.page_width_px:
            full_page = Image.new(self.original_image.mode, (self.page_width_px, self.page_height_px), (255, 255, 255))
            if self.right_to_left.get():
                # Pad on the left side for RTL (partial page is leftmost)
                full_page.paste(page_img, (self.page_width_px - page_img.size[0], 0))
            else:
                # Pad on the right side for LTR (partial page is rightmost)
                full_page.paste(page_img, (0, 0))
            page_img = full_page

        # Generate filename
        base_name = os.path.splitext(os.path.basename(self.image_path))[0]
        filename = f"{base_name}_page{page_num:03d}.png"
        output_path = os.path.join(self.output_dir, filename)

        # Save with high quality
        page_img.save(output_path, "PNG", compress_level=1)

        self.status_label.config(text=f"Exported: {filename}")

    def export_all_pages(self):
        if not self.original_image or not self.output_dir:
            messagebox.showwarning("Warning", "Please load an image and set output directory first")
            return

        total = self.pages_x
        if not messagebox.askyesno("Confirm", f"Export all {total} pages?"):
            return

        for page_num in range(1, self.pages_x + 1):
            self.export_page(page_num)
            self.status_label.config(text=f"Exporting page {page_num}/{total}...")
            self.root.update()

        self.status_label.config(text=f"Exported all {total} pages to {self.output_dir}")
        messagebox.showinfo("Complete", f"Exported {total} pages")

    def export_visible_pages(self):
        if not self.original_image or not self.output_dir:
            messagebox.showwarning("Warning", "Please load an image and set output directory first")
            return

        canvas_w = self.canvas.winfo_width()

        # Calculate visible page range
        view_left = self.pan_x
        view_right = self.pan_x + canvas_w / self.zoom

        if self.right_to_left.get():
            start_x = self.img_width - self.start_offset
            start_page = max(1, int((start_x - view_right) // self.page_width_px) + 1)
            end_page = min(self.pages_x, int((start_x - view_left) // self.page_width_px) + 1)
        else:
            start_x = self.start_offset
            start_page = max(1, int((view_left - start_x) // self.page_width_px) + 1)
            end_page = min(self.pages_x, int((view_right - start_x) // self.page_width_px) + 1)

        count = end_page - start_page + 1

        if not messagebox.askyesno("Confirm", f"Export {count} visible pages ({start_page}-{end_page})?"):
            return

        exported = 0
        for page_num in range(start_page, end_page + 1):
            self.export_page(page_num)
            exported += 1
            self.status_label.config(text=f"Exporting page {page_num} ({exported}/{count})...")
            self.root.update()

        self.status_label.config(text=f"Exported {count} visible pages to {self.output_dir}")
        messagebox.showinfo("Complete", f"Exported {count} pages")

    def show_print_dialog(self):
        """Show print options dialog."""
        if not self.original_image:
            messagebox.showwarning("Warning", "Please load an image first")
            return

        if not HAS_WIN32:
            messagebox.showerror("Error", "Printing requires pywin32.\nInstall with: pip install pywin32")
            return

        # Create print dialog
        dialog = tk.Toplevel(self.root)
        dialog.title("Print Pages")
        dialog.geometry("400x400")
        dialog.resizable(False, False)
        dialog.transient(self.root)
        dialog.grab_set()

        # Printer selection
        printer_frame = ttk.LabelFrame(dialog, text="Printer", padding=10)
        printer_frame.pack(fill=tk.X, padx=10, pady=10)

        # Get list of printers
        printers = [p[2] for p in win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS)]
        default_printer = win32print.GetDefaultPrinter()

        printer_var = tk.StringVar(value=default_printer)
        printer_combo = ttk.Combobox(printer_frame, textvariable=printer_var, values=printers, state="readonly", width=45)
        printer_combo.pack(fill=tk.X)

        # Page selection
        frame = ttk.LabelFrame(dialog, text="Pages to Print", padding=10)
        frame.pack(fill=tk.X, padx=10, pady=5)

        page_choice = tk.StringVar(value="all")

        ttk.Radiobutton(frame, text=f"All pages (1-{self.pages_x})",
                       variable=page_choice, value="all").pack(anchor=tk.W)
        ttk.Radiobutton(frame, text="Visible pages only",
                       variable=page_choice, value="visible").pack(anchor=tk.W)

        range_frame = ttk.Frame(frame)
        range_frame.pack(fill=tk.X, pady=5)
        ttk.Radiobutton(range_frame, text="Range:",
                       variable=page_choice, value="range").pack(side=tk.LEFT)
        range_from = ttk.Entry(range_frame, width=5)
        range_from.pack(side=tk.LEFT, padx=2)
        range_from.insert(0, "1")
        ttk.Label(range_frame, text="to").pack(side=tk.LEFT, padx=2)
        range_to = ttk.Entry(range_frame, width=5)
        range_to.pack(side=tk.LEFT, padx=2)
        range_to.insert(0, str(self.pages_x))

        # Print options
        opt_frame = ttk.LabelFrame(dialog, text="Options", padding=10)
        opt_frame.pack(fill=tk.X, padx=10, pady=5)

        fit_to_page = tk.BooleanVar(value=True)
        ttk.Checkbutton(opt_frame, text="Fit to page", variable=fit_to_page).pack(anchor=tk.W)

        single_sided = tk.BooleanVar(value=True)
        ttk.Checkbutton(opt_frame, text="Single-sided (no duplex)", variable=single_sided).pack(anchor=tk.W)

        ttk.Label(opt_frame, text="Orientation: Landscape (auto)",
                 foreground="gray").pack(anchor=tk.W)

        # Buttons
        btn_frame = ttk.Frame(dialog)
        btn_frame.pack(fill=tk.X, padx=10, pady=10)

        def do_print():
            choice = page_choice.get()
            if choice == "all":
                pages = list(range(1, self.pages_x + 1))
            elif choice == "visible":
                pages = self.get_visible_page_numbers()
            else:
                try:
                    start = int(range_from.get())
                    end = int(range_to.get())
                    pages = list(range(max(1, start), min(self.pages_x, end) + 1))
                except ValueError:
                    messagebox.showerror("Error", "Invalid page range")
                    return

            if not pages:
                messagebox.showwarning("Warning", "No pages to print")
                return

            selected_printer = printer_var.get()
            dialog.destroy()
            self.print_pages(pages, fit_to_page.get(), selected_printer, single_sided.get())

        ttk.Button(btn_frame, text="Print", command=do_print).pack(side=tk.RIGHT, padx=5)
        ttk.Button(btn_frame, text="Cancel", command=dialog.destroy).pack(side=tk.RIGHT)

    def get_visible_page_numbers(self):
        """Get list of page numbers currently visible."""
        canvas_w = self.canvas.winfo_width()
        view_left = self.pan_x
        view_right = self.pan_x + canvas_w / self.zoom

        if self.right_to_left.get():
            start_x = self.img_width - self.start_offset
            start_page = max(1, int((start_x - view_right) // self.page_width_px) + 1)
            end_page = min(self.pages_x, int((start_x - view_left) // self.page_width_px) + 1)
        else:
            start_x = self.start_offset
            start_page = max(1, int((view_left - start_x) // self.page_width_px) + 1)
            end_page = min(self.pages_x, int((view_right - start_x) // self.page_width_px) + 1)

        return list(range(start_page, end_page + 1))

    def print_pages(self, page_numbers, fit_to_page=True, printer_name=None, single_sided=True):
        """Print specified pages to the specified printer."""
        self.status_label.config(text="Preparing to print...")
        self.root.update()

        try:
            # Use default printer if none specified
            if not printer_name:
                printer_name = win32print.GetDefaultPrinter()

            # Open printer and set DEVMODE for this job
            hprinter = win32print.OpenPrinter(printer_name)
            try:
                # Get current DEVMODE
                devmode = win32print.GetPrinter(hprinter, 2)["pDevMode"]

                if devmode:
                    # Modify settings
                    if single_sided:
                        devmode.Duplex = 1  # DMDUP_SIMPLEX
                        devmode.Fields |= 0x00001000  # DM_DUPLEX

                    devmode.Orientation = 2  # DMORIENT_LANDSCAPE
                    devmode.Fields |= 0x00000001  # DM_ORIENTATION

                    # Apply modified DEVMODE to printer
                    # DM_IN_BUFFER = 8, DM_OUT_BUFFER = 2
                    win32print.DocumentProperties(
                        0, hprinter, printer_name, devmode, devmode, 8 | 2
                    )

            finally:
                win32print.ClosePrinter(hprinter)

            # Create device context - it should pick up the modified settings
            hdc = win32ui.CreateDC()
            hdc.CreatePrinterDC(printer_name)

            # Get printable area
            printable_width = hdc.GetDeviceCaps(win32con.HORZRES)
            printable_height = hdc.GetDeviceCaps(win32con.VERTRES)

            # Start print job
            hdc.StartDoc("Panorama Pages")

            for i, page_num in enumerate(page_numbers):
                self.status_label.config(text=f"Printing page {page_num} ({i+1}/{len(page_numbers)})...")
                self.root.update()

                # Get page image
                page_img = self.get_page_image(page_num)

                # Convert to RGB if necessary
                if page_img.mode != "RGB":
                    page_img = page_img.convert("RGB")

                # Calculate scaling to fit page
                img_w, img_h = page_img.size

                if fit_to_page:
                    # Scale to fit printable area while maintaining aspect ratio
                    scale_w = printable_width / img_w
                    scale_h = printable_height / img_h
                    scale = min(scale_w, scale_h)

                    new_w = int(img_w * scale)
                    new_h = int(img_h * scale)

                    # Center on page
                    x = (printable_width - new_w) // 2
                    y = (printable_height - new_h) // 2
                else:
                    new_w, new_h = img_w, img_h
                    x, y = 0, 0

                # Start page
                hdc.StartPage()

                # Draw image
                dib = ImageWin.Dib(page_img)
                dib.draw(hdc.GetHandleOutput(), (x, y, x + new_w, y + new_h))

                # End page
                hdc.EndPage()

            # End print job
            hdc.EndDoc()
            hdc.DeleteDC()

            self.status_label.config(text=f"Printed {len(page_numbers)} pages")
            messagebox.showinfo("Complete", f"Printed {len(page_numbers)} pages to {printer_name}")

        except Exception as e:
            self.status_label.config(text="Print failed")
            messagebox.showerror("Print Error", f"Failed to print: {e}")

    def get_page_image(self, page_num):
        """Get a page image by number (without saving to file)."""
        if self.right_to_left.get():
            start_x = self.img_width - self.start_offset
            right = start_x - (page_num - 1) * self.page_width_px
            left = max(0, right - self.page_width_px)
        else:
            start_x = self.start_offset
            left = start_x + (page_num - 1) * self.page_width_px
            right = min(left + self.page_width_px, self.img_width)

        top = 0
        bottom = self.img_height

        page_img = self.original_image.crop((left, top, right, bottom))

        # Pad if needed
        if page_img.size[0] != self.page_width_px:
            full_page = Image.new(self.original_image.mode,
                                 (self.page_width_px, self.page_height_px), (255, 255, 255))
            if self.right_to_left.get():
                full_page.paste(page_img, (self.page_width_px - page_img.size[0], 0))
            else:
                full_page.paste(page_img, (0, 0))
            page_img = full_page

        return page_img


def main():
    root = tk.Tk()
    app = PanoramaSlicer(root)
    root.mainloop()


if __name__ == "__main__":
    main()
