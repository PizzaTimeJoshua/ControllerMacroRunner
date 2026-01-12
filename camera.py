"""
Camera module for Controller Macro Runner.
Handles camera device enumeration, video capture, and display.
"""
import subprocess
import tkinter as tk
from tkinter import ttk
import json
import numpy as np
from PIL import Image, ImageTk

from utils import ffmpeg_path


def list_dshow_video_devices():
    """List DirectShow video devices using ffmpeg."""
    try:
        p = subprocess.run(
            [ffmpeg_path(), "-list_devices", "true", "-f", "dshow", "-i", "dummy"],
            capture_output=True
        )

    except FileNotFoundError:
        return []

    raw = p.stderr or b""
    try:
        text = raw.decode("mbcs", errors="replace")
    except Exception:
        text = raw.decode("utf-8", errors="replace")

    devices = []
    for line in text.splitlines():
        s = line.strip()
        if "Alternative name" in s:
            continue
        if "(video)" not in s:
            continue
        first = s.find('"')
        last = s.rfind('"')
        if first != -1 and last != -1 and last > first:
            name = s[first + 1:last].strip()
            if name and name not in devices:
                devices.append(name)
    return devices


def scale_image_to_fit(img: Image.Image, max_width: int, max_height: int) -> Image.Image:
    """
    Scale an image to fit within max_width x max_height while preserving aspect ratio.
    Returns the scaled image.
    """
    if max_width <= 0 or max_height <= 0:
        return img

    orig_w, orig_h = img.size
    if orig_w <= 0 or orig_h <= 0:
        return img

    # Calculate scale factor to fit within bounds
    scale_w = max_width / orig_w
    scale_h = max_height / orig_h
    scale = min(scale_w, scale_h)

    # Don't upscale beyond 1.0 for main window, but allow for popout/fullscreen
    # Actually, we want to allow scaling up for fullscreen, so no cap here

    new_w = max(1, int(orig_w * scale))
    new_h = max(1, int(orig_h * scale))

    if new_w == orig_w and new_h == orig_h:
        return img

    return img.resize((new_w, new_h), Image.Resampling.LANCZOS)


class CameraPopoutWindow:
    """Separate window for camera display with fullscreen support"""

    def __init__(self, app, on_close_callback):
        self.app = app
        self.on_close_callback = on_close_callback
        self.is_fullscreen = False

        # Create toplevel window
        self.window = tk.Toplevel(app.root)
        self.window.title("Camera View")
        self.window.geometry("800x600")
        self.window.configure(bg="black")

        # Create video label with black background, centered
        self.video_label = tk.Label(self.window, anchor="center", bg="black")
        self.video_label.pack(fill=tk.BOTH, expand=True)

        # Bind events
        self.video_label.bind("<Motion>", self._on_video_mouse_move)
        self.video_label.bind("<Leave>", self._on_video_mouse_leave)
        self.video_label.bind("<Button-1>", self._on_video_click_copy)
        self.video_label.bind("<Shift-Button-1>", self._on_video_click_copy_json)
        self.video_label.bind("<Double-Button-1>", self._toggle_fullscreen)

        # Bind ESC key to exit fullscreen
        self.window.bind("<Escape>", self._exit_fullscreen)

        # Handle window close
        self.window.protocol("WM_DELETE_WINDOW", self._on_window_close)

        # Coordinate display
        self.coord_var = tk.StringVar(value="x: -, y: -")
        coord_bar = ttk.Frame(self.window)
        coord_bar.pack(side=tk.BOTTOM, fill=tk.X, pady=(6, 0))
        ttk.Label(coord_bar, textvariable=self.coord_var).pack(side=tk.LEFT, padx=6)

        # Track display size and offset for coordinate mapping
        self._disp_img_w = 0
        self._disp_img_h = 0
        self._video_offset_x = 0  # Offset of video within the display area
        self._video_offset_y = 0
        self._last_video_xy = None

    def _toggle_fullscreen(self, event=None):
        """Toggle between windowed and fullscreen mode"""
        if self.is_fullscreen:
            self._exit_fullscreen()
        else:
            self._enter_fullscreen()

    def _enter_fullscreen(self, event=None):
        """Enter fullscreen mode"""
        self.is_fullscreen = True
        self.window.attributes("-fullscreen", True)
        self.app.set_status("Camera fullscreen (press ESC to exit)")

    def _exit_fullscreen(self, event=None):
        """Exit fullscreen mode"""
        if self.is_fullscreen:
            self.is_fullscreen = False
            self.window.attributes("-fullscreen", False)
            self.app.set_status("Camera windowed")

    def _on_window_close(self):
        """Handle window close - return to embedded mode"""
        self._exit_fullscreen()
        self.on_close_callback()

    def update_frame(self, pil_img):
        """Update the video display with a new frame, scaling to fit window and centering"""
        if pil_img is None:
            return

        # Get available size for the video (window size minus coord bar)
        self.window.update_idletasks()
        available_w = self.video_label.winfo_width()
        available_h = self.video_label.winfo_height()

        # Fallback to window size if label size not yet available
        if available_w <= 1 or available_h <= 1:
            available_w = self.window.winfo_width()
            available_h = self.window.winfo_height() - 30  # Approximate coord bar height

        # Scale image to fit while maintaining aspect ratio
        if available_w > 1 and available_h > 1:
            scaled_img = scale_image_to_fit(pil_img, available_w, available_h)
        else:
            scaled_img = pil_img

        # Calculate offset for centering (used for coordinate mapping)
        scaled_w, scaled_h = scaled_img.size
        self._video_offset_x = (available_w - scaled_w) // 2
        self._video_offset_y = (available_h - scaled_h) // 2

        # Convert to PhotoImage and display
        tk_img = ImageTk.PhotoImage(scaled_img)
        self._disp_img_w = tk_img.width()
        self._disp_img_h = tk_img.height()
        self.video_label.imgtk = tk_img
        self.video_label.configure(image=tk_img)

    def _event_to_frame_xy(self, event):
        """Convert mouse event to frame coordinates"""
        with self.app.frame_lock:
            frame = self.app.latest_frame_bgr
            if frame is None:
                return None
            fh, fw, _ = frame.shape

        iw = self._disp_img_w or fw
        ih = self._disp_img_h or fh

        # Account for centering offset
        off_x = self._video_offset_x
        off_y = self._video_offset_y
        x_img = int(event.x) - off_x
        y_img = int(event.y) - off_y

        if not (0 <= x_img < iw and 0 <= y_img < ih):
            return None

        x = int(x_img * fw / iw)
        y = int(y_img * fh / ih)

        if 0 <= x < fw and 0 <= y < fh:
            return (x, y)
        return None

    def _on_video_mouse_move(self, event):
        """Handle mouse movement over video"""
        xy = self._event_to_frame_xy(event)
        if xy is None:
            self._last_video_xy = None
            self.coord_var.set("x: -, y: -")
            return
        x, y = xy
        self._last_video_xy = (x, y)
        self.coord_var.set(f"x: {x}, y: {y}")

    def _on_video_mouse_leave(self, event):
        """Handle mouse leaving video area"""
        self._last_video_xy = None
        self.coord_var.set("x: -, y: -")

    def _on_video_click_copy(self, event):
        """Copy coordinates on click"""
        xy = self._event_to_frame_xy(event) or self._last_video_xy
        if xy is None:
            self.app.set_status("No coords to copy.")
            return
        x, y = xy
        s = f"{x},{y}"
        self.app._copy_to_clipboard(s)
        self.app.set_status(f"Copied coords: {s}")

    def _on_video_click_copy_json(self, event):
        """Copy coordinates as JSON on Shift+click"""
        xy = self._event_to_frame_xy(event) or self._last_video_xy
        if xy is None:
            self.app.set_status("No coords to copy.")
            return
        x, y = xy
        s = json.dumps({"x": x, "y": y})
        self.app._copy_to_clipboard(s)
        self.app.set_status(f"Copied coords JSON: {s}")

    def close(self):
        """Close the window"""
        try:
            self.window.destroy()
        except:
            pass


class RegionSelectorWindow:
    """
    Window for selecting a region on the camera feed.
    User clicks and drags to draw a rectangle, then confirms the selection.
    """

    def __init__(self, app, on_select_callback, initial_region=None, on_close_callback=None):
        """
        Args:
            app: Main application instance (for frame access)
            on_select_callback: Called with (x, y, width, height) when confirmed
            initial_region: Optional tuple (x, y, width, height) to show initially
            on_close_callback: Optional callback called when window closes (for any reason)
        """
        self.app = app
        self.on_select_callback = on_select_callback
        self.on_close_callback = on_close_callback
        self.result = None

        # Selection state
        self.start_x = None
        self.start_y = None
        self.current_rect = initial_region  # (x, y, w, h) in frame coords
        self.dragging = False

        # Create window
        self.window = tk.Toplevel(app.root)
        self.window.title("Select Region - Click and drag to select area")
        self.window.geometry("800x600")
        self.window.configure(bg="black")
        self.window.transient(app.root)
        self.window.grab_set()

        # Create canvas for video and drawing
        self.canvas = tk.Canvas(self.window, bg="black", highlightthickness=0)
        self.canvas.pack(fill=tk.BOTH, expand=True)

        # Bottom bar with info and buttons
        bottom = ttk.Frame(self.window)
        bottom.pack(side=tk.BOTTOM, fill=tk.X, pady=6, padx=6)

        self.info_var = tk.StringVar(value="Click and drag to select region")
        ttk.Label(bottom, textvariable=self.info_var).pack(side=tk.LEFT)

        ttk.Button(bottom, text="Cancel", command=self._cancel).pack(side=tk.RIGHT, padx=(6, 0))
        ttk.Button(bottom, text="Confirm", command=self._confirm).pack(side=tk.RIGHT)
        ttk.Button(bottom, text="Clear", command=self._clear).pack(side=tk.RIGHT, padx=(0, 6))

        # Bind mouse events
        self.canvas.bind("<ButtonPress-1>", self._on_mouse_down)
        self.canvas.bind("<B1-Motion>", self._on_mouse_drag)
        self.canvas.bind("<ButtonRelease-1>", self._on_mouse_up)
        self.canvas.bind("<Motion>", self._on_mouse_move)

        # Handle window close
        self.window.protocol("WM_DELETE_WINDOW", self._cancel)

        # Display size tracking
        self._disp_img_w = 0
        self._disp_img_h = 0
        self._img_offset_x = 0
        self._img_offset_y = 0

        # Start frame updates
        self._update_loop()

    def _update_loop(self):
        """Update the display with current frame and selection overlay"""
        if not self.window.winfo_exists():
            return

        self._update_frame()
        self.window.after(30, self._update_loop)

    def _update_frame(self):
        """Draw current frame with selection rectangle overlay"""
        with self.app.frame_lock:
            frame = self.app.latest_frame_bgr
            if frame is None:
                return
            frame = frame.copy()

        # Convert BGR to RGB
        rgb = frame[:, :, ::-1]
        img = Image.fromarray(rgb)

        # Get canvas size
        self.canvas.update_idletasks()
        canvas_w = self.canvas.winfo_width()
        canvas_h = self.canvas.winfo_height()

        if canvas_w <= 1 or canvas_h <= 1:
            return

        # Scale image to fit
        scaled_img = scale_image_to_fit(img, canvas_w, canvas_h)
        scaled_w, scaled_h = scaled_img.size

        # Calculate offset for centering
        self._img_offset_x = (canvas_w - scaled_w) // 2
        self._img_offset_y = (canvas_h - scaled_h) // 2
        self._disp_img_w = scaled_w
        self._disp_img_h = scaled_h

        # Store frame dimensions for coordinate conversion
        self._frame_w = frame.shape[1]
        self._frame_h = frame.shape[0]

        # Convert to PhotoImage
        tk_img = ImageTk.PhotoImage(scaled_img)
        self.canvas.imgtk = tk_img

        # Clear and redraw
        self.canvas.delete("all")
        self.canvas.create_image(
            self._img_offset_x, self._img_offset_y,
            anchor="nw", image=tk_img
        )

        # Draw selection rectangle if we have one
        if self.current_rect:
            x, y, w, h = self.current_rect
            # Convert frame coords to canvas coords
            cx1, cy1 = self._frame_to_canvas(x, y)
            cx2, cy2 = self._frame_to_canvas(x + w, y + h)
            self.canvas.create_rectangle(
                cx1, cy1, cx2, cy2,
                outline="#00ff00", width=2, dash=(4, 4)
            )
            # Draw corner handles
            handle_size = 6
            for hx, hy in [(cx1, cy1), (cx2, cy1), (cx1, cy2), (cx2, cy2)]:
                self.canvas.create_rectangle(
                    hx - handle_size, hy - handle_size,
                    hx + handle_size, hy + handle_size,
                    fill="#00ff00", outline="#ffffff"
                )

    def _frame_to_canvas(self, fx, fy):
        """Convert frame coordinates to canvas coordinates"""
        if self._disp_img_w <= 0 or self._frame_w <= 0:
            return fx, fy
        cx = self._img_offset_x + int(fx * self._disp_img_w / self._frame_w)
        cy = self._img_offset_y + int(fy * self._disp_img_h / self._frame_h)
        return cx, cy

    def _canvas_to_frame(self, cx, cy):
        """Convert canvas coordinates to frame coordinates"""
        if self._disp_img_w <= 0 or self._frame_w <= 0:
            return None

        # Remove offset
        ix = cx - self._img_offset_x
        iy = cy - self._img_offset_y

        # Check if within image bounds
        if not (0 <= ix < self._disp_img_w and 0 <= iy < self._disp_img_h):
            return None

        # Scale to frame coordinates
        fx = int(ix * self._frame_w / self._disp_img_w)
        fy = int(iy * self._frame_h / self._disp_img_h)

        # Clamp to frame bounds
        fx = max(0, min(fx, self._frame_w - 1))
        fy = max(0, min(fy, self._frame_h - 1))

        return fx, fy

    def _on_mouse_down(self, event):
        """Start rectangle selection"""
        coords = self._canvas_to_frame(event.x, event.y)
        if coords is None:
            return

        self.start_x, self.start_y = coords
        self.dragging = True
        self.current_rect = None

    def _on_mouse_drag(self, event):
        """Update rectangle during drag"""
        if not self.dragging or self.start_x is None:
            return

        coords = self._canvas_to_frame(event.x, event.y)
        if coords is None:
            return

        end_x, end_y = coords

        # Calculate rectangle (handle any drag direction)
        x = min(self.start_x, end_x)
        y = min(self.start_y, end_y)
        w = abs(end_x - self.start_x)
        h = abs(end_y - self.start_y)

        # Minimum size
        w = max(1, w)
        h = max(1, h)

        self.current_rect = (x, y, w, h)
        self.info_var.set(f"Region: ({x}, {y}) {w}x{h}")

    def _on_mouse_up(self, event):
        """Finish rectangle selection"""
        if self.dragging:
            self._on_mouse_drag(event)  # Finalize position
            self.dragging = False

            if self.current_rect and self.current_rect[2] > 0 and self.current_rect[3] > 0:
                x, y, w, h = self.current_rect
                self.info_var.set(f"Region: ({x}, {y}) {w}x{h} - Click Confirm to use")

    def _on_mouse_move(self, event):
        """Show coordinates while hovering"""
        if self.dragging:
            return

        coords = self._canvas_to_frame(event.x, event.y)
        if coords is None:
            if not self.current_rect:
                self.info_var.set("Click and drag to select region")
            return

        fx, fy = coords
        if self.current_rect:
            x, y, w, h = self.current_rect
            self.info_var.set(f"Region: ({x}, {y}) {w}x{h} | Cursor: ({fx}, {fy})")
        else:
            self.info_var.set(f"Cursor: ({fx}, {fy}) - Click and drag to select")

    def _clear(self):
        """Clear the current selection"""
        self.current_rect = None
        self.info_var.set("Click and drag to select region")

    def _confirm(self):
        """Confirm selection and close"""
        if self.current_rect and self.current_rect[2] > 0 and self.current_rect[3] > 0:
            self.result = self.current_rect
            if self.on_select_callback:
                self.on_select_callback(*self.current_rect)
            self._close_window()
        else:
            self.info_var.set("Please select a region first")

    def _cancel(self):
        """Cancel and close"""
        self.result = None
        self._close_window()

    def _close_window(self):
        """Close the window and call the close callback"""
        self.window.destroy()
        if self.on_close_callback:
            self.on_close_callback()


class ColorPickerWindow:
    """
    Window for picking a color from the camera feed.
    User clicks on a pixel to select its color.
    Also allows picking X,Y coordinates for the find_color command.
    """

    def __init__(self, app, on_select_callback, initial_x=None, initial_y=None, on_close_callback=None):
        """
        Args:
            app: Main application instance (for frame access)
            on_select_callback: Called with (x, y, r, g, b) when confirmed
            initial_x: Optional initial X coordinate
            initial_y: Optional initial Y coordinate
            on_close_callback: Optional callback called when window closes (for any reason)
        """
        self.app = app
        self.on_select_callback = on_select_callback
        self.on_close_callback = on_close_callback
        self.result = None

        # Selection state
        self.selected_x = initial_x
        self.selected_y = initial_y
        self.selected_rgb = None  # (r, g, b)
        self.hover_x = None
        self.hover_y = None
        self.hover_rgb = None

        # Create window
        self.window = tk.Toplevel(app.root)
        self.window.title("Pick Color - Click on camera to select color and position")
        self.window.geometry("800x600")
        self.window.configure(bg="black")
        self.window.transient(app.root)
        self.window.grab_set()

        # Main container
        main_frame = ttk.Frame(self.window)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Create canvas for video
        self.canvas = tk.Canvas(main_frame, bg="black", highlightthickness=0)
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Right side panel for color info
        right_panel = ttk.Frame(main_frame, width=200)
        right_panel.pack(side=tk.RIGHT, fill=tk.Y, padx=6, pady=6)
        right_panel.pack_propagate(False)

        # Hover color section
        ttk.Label(right_panel, text="Hover Color:", font=("", 10, "bold")).pack(anchor="w", pady=(0, 4))
        self.hover_swatch = tk.Canvas(right_panel, width=180, height=60, bg="#808080", highlightthickness=1, highlightbackground="#555")
        self.hover_swatch.pack(pady=(0, 4))
        self.hover_info_var = tk.StringVar(value="Move cursor over image")
        ttk.Label(right_panel, textvariable=self.hover_info_var, wraplength=180).pack(anchor="w")

        ttk.Separator(right_panel, orient="horizontal").pack(fill="x", pady=10)

        # Selected color section
        ttk.Label(right_panel, text="Selected Color:", font=("", 10, "bold")).pack(anchor="w", pady=(0, 4))
        self.selected_swatch = tk.Canvas(right_panel, width=180, height=60, bg="#808080", highlightthickness=1, highlightbackground="#555")
        self.selected_swatch.pack(pady=(0, 4))
        self.selected_info_var = tk.StringVar(value="Click to select")
        ttk.Label(right_panel, textvariable=self.selected_info_var, wraplength=180).pack(anchor="w")

        # Bottom bar with buttons
        bottom = ttk.Frame(self.window)
        bottom.pack(side=tk.BOTTOM, fill=tk.X, pady=6, padx=6)

        self.info_var = tk.StringVar(value="Click on camera to pick a color")
        ttk.Label(bottom, textvariable=self.info_var).pack(side=tk.LEFT)

        ttk.Button(bottom, text="Cancel", command=self._cancel).pack(side=tk.RIGHT, padx=(6, 0))
        ttk.Button(bottom, text="Confirm", command=self._confirm).pack(side=tk.RIGHT)
        ttk.Button(bottom, text="Clear", command=self._clear).pack(side=tk.RIGHT, padx=(0, 6))

        # Bind mouse events
        self.canvas.bind("<Motion>", self._on_mouse_move)
        self.canvas.bind("<Leave>", self._on_mouse_leave)
        self.canvas.bind("<ButtonPress-1>", self._on_click)

        # Handle window close
        self.window.protocol("WM_DELETE_WINDOW", self._cancel)

        # Display size tracking
        self._disp_img_w = 0
        self._disp_img_h = 0
        self._img_offset_x = 0
        self._img_offset_y = 0
        self._frame_w = 0
        self._frame_h = 0

        # Initialize with existing selection if provided
        if initial_x is not None and initial_y is not None:
            self._update_selection_from_coords(initial_x, initial_y)

        # Start frame updates
        self._update_loop()

    def _update_loop(self):
        """Update the display with current frame"""
        if not self.window.winfo_exists():
            return

        self._update_frame()
        self.window.after(30, self._update_loop)

    def _update_frame(self):
        """Draw current frame with crosshair on selected point"""
        with self.app.frame_lock:
            frame = self.app.latest_frame_bgr
            if frame is None:
                return
            frame = frame.copy()

        self._frame_w = frame.shape[1]
        self._frame_h = frame.shape[0]

        # Convert BGR to RGB
        rgb = frame[:, :, ::-1]
        img = Image.fromarray(rgb)

        # Get canvas size
        self.canvas.update_idletasks()
        canvas_w = self.canvas.winfo_width()
        canvas_h = self.canvas.winfo_height()

        if canvas_w <= 1 or canvas_h <= 1:
            return

        # Scale image to fit
        scaled_img = scale_image_to_fit(img, canvas_w, canvas_h)
        scaled_w, scaled_h = scaled_img.size

        # Calculate offset for centering
        self._img_offset_x = (canvas_w - scaled_w) // 2
        self._img_offset_y = (canvas_h - scaled_h) // 2
        self._disp_img_w = scaled_w
        self._disp_img_h = scaled_h

        # Convert to PhotoImage
        tk_img = ImageTk.PhotoImage(scaled_img)
        self.canvas.imgtk = tk_img

        # Clear and redraw
        self.canvas.delete("all")
        self.canvas.create_image(
            self._img_offset_x, self._img_offset_y,
            anchor="nw", image=tk_img
        )

        # Draw crosshair on selected point
        if self.selected_x is not None and self.selected_y is not None:
            cx, cy = self._frame_to_canvas(self.selected_x, self.selected_y)
            line_len = 15
            # Vertical line
            self.canvas.create_line(cx, cy - line_len, cx, cy + line_len, fill="#00ff00", width=2)
            # Horizontal line
            self.canvas.create_line(cx - line_len, cy, cx + line_len, cy, fill="#00ff00", width=2)
            # Center circle
            self.canvas.create_oval(cx - 5, cy - 5, cx + 5, cy + 5, outline="#00ff00", width=2)

        # Update hover color if we have coordinates
        if self.hover_x is not None and self.hover_y is not None:
            if 0 <= self.hover_x < self._frame_w and 0 <= self.hover_y < self._frame_h:
                b, g, r = frame[self.hover_y, self.hover_x]
                self.hover_rgb = (int(r), int(g), int(b))
                hex_color = f"#{r:02x}{g:02x}{b:02x}"
                self.hover_swatch.configure(bg=hex_color)
                self.hover_info_var.set(f"({self.hover_x}, {self.hover_y})\nRGB: {r}, {g}, {b}")

    def _frame_to_canvas(self, fx, fy):
        """Convert frame coordinates to canvas coordinates"""
        if self._disp_img_w <= 0 or self._frame_w <= 0:
            return fx, fy
        cx = self._img_offset_x + int(fx * self._disp_img_w / self._frame_w)
        cy = self._img_offset_y + int(fy * self._disp_img_h / self._frame_h)
        return cx, cy

    def _canvas_to_frame(self, cx, cy):
        """Convert canvas coordinates to frame coordinates"""
        if self._disp_img_w <= 0 or self._frame_w <= 0:
            return None

        # Remove offset
        ix = cx - self._img_offset_x
        iy = cy - self._img_offset_y

        # Check if within image bounds
        if not (0 <= ix < self._disp_img_w and 0 <= iy < self._disp_img_h):
            return None

        # Scale to frame coordinates
        fx = int(ix * self._frame_w / self._disp_img_w)
        fy = int(iy * self._frame_h / self._disp_img_h)

        # Clamp to frame bounds
        fx = max(0, min(fx, self._frame_w - 1))
        fy = max(0, min(fy, self._frame_h - 1))

        return fx, fy

    def _on_mouse_move(self, event):
        """Update hover info"""
        coords = self._canvas_to_frame(event.x, event.y)
        if coords is None:
            self.hover_x = None
            self.hover_y = None
            self.hover_info_var.set("Move cursor over image")
            self.hover_swatch.configure(bg="#808080")
            return

        self.hover_x, self.hover_y = coords

    def _on_mouse_leave(self, event):
        """Clear hover info"""
        self.hover_x = None
        self.hover_y = None
        self.hover_info_var.set("Move cursor over image")
        self.hover_swatch.configure(bg="#808080")

    def _on_click(self, event):
        """Select color at clicked point"""
        coords = self._canvas_to_frame(event.x, event.y)
        if coords is None:
            return

        self._update_selection_from_coords(coords[0], coords[1])

    def _update_selection_from_coords(self, x, y):
        """Update selection based on coordinates"""
        self.selected_x = x
        self.selected_y = y

        # Get color from current frame
        with self.app.frame_lock:
            frame = self.app.latest_frame_bgr
            if frame is not None and 0 <= y < frame.shape[0] and 0 <= x < frame.shape[1]:
                b, g, r = frame[y, x]
                self.selected_rgb = (int(r), int(g), int(b))
                hex_color = f"#{r:02x}{g:02x}{b:02x}"
                self.selected_swatch.configure(bg=hex_color)
                self.selected_info_var.set(f"({x}, {y})\nRGB: {r}, {g}, {b}")
                self.info_var.set(f"Selected: ({x}, {y}) RGB({r}, {g}, {b}) - Click Confirm to use")

    def _clear(self):
        """Clear the current selection"""
        self.selected_x = None
        self.selected_y = None
        self.selected_rgb = None
        self.selected_swatch.configure(bg="#808080")
        self.selected_info_var.set("Click to select")
        self.info_var.set("Click on camera to pick a color")

    def _confirm(self):
        """Confirm selection and close"""
        if self.selected_x is not None and self.selected_y is not None and self.selected_rgb is not None:
            self.result = (self.selected_x, self.selected_y, *self.selected_rgb)
            if self.on_select_callback:
                self.on_select_callback(self.selected_x, self.selected_y, *self.selected_rgb)
            self._close_window()
        else:
            self.info_var.set("Please select a color first")

    def _cancel(self):
        """Cancel and close"""
        self.result = None
        self._close_window()

    def _close_window(self):
        """Close the window and call the close callback"""
        self.window.destroy()
        if self.on_close_callback:
            self.on_close_callback()


class AreaColorPickerWindow:
    """
    Window for selecting an area and picking its average color from the camera feed.
    Combines region selection with color detection - shows average color of selected area.
    """

    def __init__(self, app, on_select_callback, initial_region=None, initial_rgb=None, on_close_callback=None):
        """
        Args:
            app: Main application instance (for frame access)
            on_select_callback: Called with (x, y, width, height, r, g, b) when confirmed
            initial_region: Optional tuple (x, y, width, height) to show initially
            initial_rgb: Optional tuple (r, g, b) for initial target color
            on_close_callback: Optional callback called when window closes (for any reason)
        """
        self.app = app
        self.on_select_callback = on_select_callback
        self.on_close_callback = on_close_callback
        self.result = None

        # Selection state
        self.start_x = None
        self.start_y = None
        self.current_rect = initial_region  # (x, y, w, h) in frame coords
        self.dragging = False
        self.avg_rgb = None  # Current average color of selected area
        self.target_rgb = initial_rgb or [255, 0, 0]  # Target color to compare against

        # Create window
        self.window = tk.Toplevel(app.root)
        self.window.title("Select Area & Color - Click and drag to select area")
        self.window.geometry("1000x600")
        self.window.configure(bg="black")
        self.window.transient(app.root)
        self.window.grab_set()

        # Main container
        main_frame = ttk.Frame(self.window)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Create canvas for video and drawing
        self.canvas = tk.Canvas(main_frame, bg="black", highlightthickness=0)
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Right side panel for color info
        right_panel = ttk.Frame(main_frame, width=250)
        right_panel.pack(side=tk.RIGHT, fill=tk.Y, padx=6, pady=6)
        right_panel.pack_propagate(False)

        # Area average color section
        ttk.Label(right_panel, text="Area Average Color:", font=("", 10, "bold")).pack(anchor="w", pady=(0, 4))
        self.avg_swatch = tk.Canvas(right_panel, width=230, height=80, bg="#808080",
                                     highlightthickness=1, highlightbackground="#555")
        self.avg_swatch.pack(pady=(0, 4))
        self.avg_info_var = tk.StringVar(value="Select an area to see average color")
        ttk.Label(right_panel, textvariable=self.avg_info_var, wraplength=230).pack(anchor="w")

        # Copy button to copy average to target
        ttk.Button(right_panel, text="↓ Copy Average to Target",
                   command=self._copy_avg_to_target).pack(pady=(4, 0))

        ttk.Separator(right_panel, orient="horizontal").pack(fill="x", pady=10)

        # Target color section
        ttk.Label(right_panel, text="Target Color:", font=("", 10, "bold")).pack(anchor="w", pady=(0, 4))
        self.target_swatch = tk.Canvas(right_panel, width=230, height=80, bg="#ff0000",
                                        highlightthickness=1, highlightbackground="#555")
        self.target_swatch.pack(pady=(0, 4))
        self.target_info_var = tk.StringVar(value=f"RGB: {self.target_rgb[0]}, {self.target_rgb[1]}, {self.target_rgb[2]}")
        ttk.Label(right_panel, textvariable=self.target_info_var, wraplength=230).pack(anchor="w")

        # Update target swatch with initial color
        self._update_target_swatch()

        # Bottom bar with info and buttons
        bottom = ttk.Frame(self.window)
        bottom.pack(side=tk.BOTTOM, fill=tk.X, pady=6, padx=6)

        self.info_var = tk.StringVar(value="Click and drag to select region")
        ttk.Label(bottom, textvariable=self.info_var).pack(side=tk.LEFT)

        ttk.Button(bottom, text="Cancel", command=self._cancel).pack(side=tk.RIGHT, padx=(6, 0))
        ttk.Button(bottom, text="Confirm", command=self._confirm).pack(side=tk.RIGHT)
        ttk.Button(bottom, text="Clear", command=self._clear).pack(side=tk.RIGHT, padx=(0, 6))

        # Bind mouse events
        self.canvas.bind("<ButtonPress-1>", self._on_mouse_down)
        self.canvas.bind("<B1-Motion>", self._on_mouse_drag)
        self.canvas.bind("<ButtonRelease-1>", self._on_mouse_up)
        self.canvas.bind("<Motion>", self._on_mouse_move)

        # Handle window close
        self.window.protocol("WM_DELETE_WINDOW", self._cancel)

        # Display size tracking
        self._disp_img_w = 0
        self._disp_img_h = 0
        self._img_offset_x = 0
        self._img_offset_y = 0
        self._frame_w = 0
        self._frame_h = 0

        # Start frame updates
        self._update_loop()

    def _update_loop(self):
        """Update the display with current frame and selection overlay"""
        if not self.window.winfo_exists():
            return

        self._update_frame()
        self.window.after(30, self._update_loop)

    def _update_frame(self):
        """Draw current frame with selection rectangle overlay and calculate average color"""
        with self.app.frame_lock:
            frame = self.app.latest_frame_bgr
            if frame is None:
                return
            frame = frame.copy()

        self._frame_w = frame.shape[1]
        self._frame_h = frame.shape[0]

        # Convert BGR to RGB
        rgb = frame[:, :, ::-1]
        img = Image.fromarray(rgb)

        # Get canvas size
        self.canvas.update_idletasks()
        canvas_w = self.canvas.winfo_width()
        canvas_h = self.canvas.winfo_height()

        if canvas_w <= 1 or canvas_h <= 1:
            return

        # Scale image to fit
        scaled_img = scale_image_to_fit(img, canvas_w, canvas_h)
        scaled_w, scaled_h = scaled_img.size

        # Calculate offset for centering
        self._img_offset_x = (canvas_w - scaled_w) // 2
        self._img_offset_y = (canvas_h - scaled_h) // 2
        self._disp_img_w = scaled_w
        self._disp_img_h = scaled_h

        # Convert to PhotoImage
        tk_img = ImageTk.PhotoImage(scaled_img)
        self.canvas.imgtk = tk_img

        # Clear and redraw
        self.canvas.delete("all")
        self.canvas.create_image(
            self._img_offset_x, self._img_offset_y,
            anchor="nw", image=tk_img
        )

        # Draw selection rectangle if we have one
        if self.current_rect:
            x, y, w, h = self.current_rect
            # Convert frame coords to canvas coords
            cx1, cy1 = self._frame_to_canvas(x, y)
            cx2, cy2 = self._frame_to_canvas(x + w, y + h)
            self.canvas.create_rectangle(
                cx1, cy1, cx2, cy2,
                outline="#00ff00", width=2, dash=(4, 4)
            )
            # Draw corner handles
            handle_size = 6
            for hx, hy in [(cx1, cy1), (cx2, cy1), (cx1, cy2), (cx2, cy2)]:
                self.canvas.create_rectangle(
                    hx - handle_size, hy - handle_size,
                    hx + handle_size, hy + handle_size,
                    fill="#00ff00", outline="#ffffff"
                )

            # Calculate and display average color for selected region
            self._update_avg_color(frame, x, y, w, h)

    def _update_avg_color(self, frame, x, y, w, h):
        """Calculate and display the average color of the selected region."""
        # Clamp region to frame bounds
        x = max(0, min(x, self._frame_w - 1))
        y = max(0, min(y, self._frame_h - 1))
        x2 = max(x + 1, min(x + w, self._frame_w))
        y2 = max(y + 1, min(y + h, self._frame_h))

        # Extract region (BGR)
        region_bgr = frame[y:y2, x:x2]

        if region_bgr.size == 0:
            self.avg_rgb = None
            self.avg_swatch.configure(bg="#808080")
            self.avg_info_var.set("Region is empty")
            return

        # Calculate average color
        avg_b = float(np.mean(region_bgr[:, :, 0]))
        avg_g = float(np.mean(region_bgr[:, :, 1]))
        avg_r = float(np.mean(region_bgr[:, :, 2]))

        self.avg_rgb = (int(avg_r), int(avg_g), int(avg_b))
        r, g, b = self.avg_rgb

        # Update swatch
        hex_color = f"#{r:02x}{g:02x}{b:02x}"
        self.avg_swatch.configure(bg=hex_color)
        self.avg_info_var.set(f"Region: ({x},{y}) {w}x{h}\nAvg RGB: {r}, {g}, {b}")

    def _update_target_swatch(self):
        """Update the target color swatch."""
        r, g, b = self.target_rgb
        hex_color = f"#{r:02x}{g:02x}{b:02x}"
        self.target_swatch.configure(bg=hex_color)
        self.target_info_var.set(f"RGB: {r}, {g}, {b}")

    def _copy_avg_to_target(self):
        """Copy the average color to the target color."""
        if self.avg_rgb is None:
            self.info_var.set("Select an area first to copy its average color")
            return

        self.target_rgb = list(self.avg_rgb)
        self._update_target_swatch()
        self.info_var.set(f"Copied average RGB({self.avg_rgb[0]}, {self.avg_rgb[1]}, {self.avg_rgb[2]}) to target")

    def _frame_to_canvas(self, fx, fy):
        """Convert frame coordinates to canvas coordinates"""
        if self._disp_img_w <= 0 or self._frame_w <= 0:
            return fx, fy
        cx = self._img_offset_x + int(fx * self._disp_img_w / self._frame_w)
        cy = self._img_offset_y + int(fy * self._disp_img_h / self._frame_h)
        return cx, cy

    def _canvas_to_frame(self, cx, cy):
        """Convert canvas coordinates to frame coordinates"""
        if self._disp_img_w <= 0 or self._frame_w <= 0:
            return None

        # Remove offset
        ix = cx - self._img_offset_x
        iy = cy - self._img_offset_y

        # Check if within image bounds
        if not (0 <= ix < self._disp_img_w and 0 <= iy < self._disp_img_h):
            return None

        # Scale to frame coordinates
        fx = int(ix * self._frame_w / self._disp_img_w)
        fy = int(iy * self._frame_h / self._disp_img_h)

        # Clamp to frame bounds
        fx = max(0, min(fx, self._frame_w - 1))
        fy = max(0, min(fy, self._frame_h - 1))

        return fx, fy

    def _on_mouse_down(self, event):
        """Start rectangle selection"""
        coords = self._canvas_to_frame(event.x, event.y)
        if coords is None:
            return

        self.start_x, self.start_y = coords
        self.dragging = True
        self.current_rect = None

    def _on_mouse_drag(self, event):
        """Update rectangle during drag"""
        if not self.dragging or self.start_x is None:
            return

        coords = self._canvas_to_frame(event.x, event.y)
        if coords is None:
            return

        end_x, end_y = coords

        # Calculate rectangle (handle any drag direction)
        x = min(self.start_x, end_x)
        y = min(self.start_y, end_y)
        w = abs(end_x - self.start_x)
        h = abs(end_y - self.start_y)

        # Minimum size
        w = max(1, w)
        h = max(1, h)

        self.current_rect = (x, y, w, h)
        self.info_var.set(f"Region: ({x}, {y}) {w}x{h}")

    def _on_mouse_up(self, event):
        """Finish rectangle selection"""
        if self.dragging:
            self._on_mouse_drag(event)  # Finalize position
            self.dragging = False

            if self.current_rect and self.current_rect[2] > 0 and self.current_rect[3] > 0:
                x, y, w, h = self.current_rect
                self.info_var.set(f"Region: ({x}, {y}) {w}x{h} - Click Confirm to use")

    def _on_mouse_move(self, event):
        """Show coordinates while hovering"""
        if self.dragging:
            return

        coords = self._canvas_to_frame(event.x, event.y)
        if coords is None:
            if not self.current_rect:
                self.info_var.set("Click and drag to select region")
            return

        fx, fy = coords
        if self.current_rect:
            x, y, w, h = self.current_rect
            self.info_var.set(f"Region: ({x}, {y}) {w}x{h} | Cursor: ({fx}, {fy})")
        else:
            self.info_var.set(f"Cursor: ({fx}, {fy}) - Click and drag to select")

    def _clear(self):
        """Clear the current selection"""
        self.current_rect = None
        self.avg_rgb = None
        self.avg_swatch.configure(bg="#808080")
        self.avg_info_var.set("Select an area to see average color")
        self.info_var.set("Click and drag to select region")

    def _confirm(self):
        """Confirm selection and close"""
        if self.current_rect and self.current_rect[2] > 0 and self.current_rect[3] > 0:
            x, y, w, h = self.current_rect
            r, g, b = self.target_rgb
            self.result = (x, y, w, h, r, g, b)
            if self.on_select_callback:
                self.on_select_callback(x, y, w, h, r, g, b)
            self._close_window()
        else:
            self.info_var.set("Please select a region first")

    def _cancel(self):
        """Cancel and close"""
        self.result = None
        self._close_window()

    def _close_window(self):
        """Close the window and call the close callback"""
        self.window.destroy()
        if self.on_close_callback:
            self.on_close_callback()
