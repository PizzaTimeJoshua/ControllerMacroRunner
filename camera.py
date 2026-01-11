"""
Camera module for Controller Macro Runner.
Handles camera device enumeration, video capture, and display.
"""
import subprocess
import tkinter as tk
from tkinter import ttk
import json
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

    def __init__(self, app, on_select_callback, initial_region=None):
        """
        Args:
            app: Main application instance (for frame access)
            on_select_callback: Called with (x, y, width, height) when confirmed
            initial_region: Optional tuple (x, y, width, height) to show initially
        """
        self.app = app
        self.on_select_callback = on_select_callback
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
            self.window.destroy()
        else:
            self.info_var.set("Please select a region first")

    def _cancel(self):
        """Cancel and close"""
        self.result = None
        self.window.destroy()
