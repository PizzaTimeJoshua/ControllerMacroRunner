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
