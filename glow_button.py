import tkinter as tk
import customtkinter as ctk


class GlowButton(ctk.CTkFrame):
    """Button with animated gradient border on hover."""

    def __init__(self, master, text: str, command=None, width: int = 140, height: int = 32, **kwargs) -> None:
        border = 1
        super().__init__(master, width=width + border * 2, height=height + border * 2, fg_color="transparent")
        self._border = border
        self._span = width + border * 2
        self._offset = 0
        self._animating = False

        # Canvas for gradient border
        self.canvas = tk.Canvas(
            self,
            width=self._span,
            height=height + border * 2,
            highlightthickness=0,
            bd=0,
        )
        self.canvas.place(x=0, y=0)

        # Create gradient image (#ff0000 -> #ff00c8)
        self.gradient_width = self._span * 2
        self.gradient_image = tk.PhotoImage(width=self.gradient_width, height=height + border * 2)
        for x in range(self.gradient_width):
            b = int(200 * ((x % self._span) / (self._span - 1)))
            color = f"#ff00{b:02x}"
            for y in range(height + border * 2):
                self.gradient_image.put(color, (x, y))
        self.image_id = self.canvas.create_image(0, 0, image=self.gradient_image, anchor="nw")
        self.canvas.itemconfigure(self.image_id, state="hidden")

        # Underlying CTkButton
        self.button = ctk.CTkButton(
            self,
            text="",
            command=command,
            width=width,
            height=height,
            corner_radius=10,
            **kwargs,
        )
        self.button.place(x=border, y=border)

        self._text = text
        self._font = kwargs.get("font", self.button.cget("font"))
        self._text_color = self._normalize_color(
            self._resolve_single_color(kwargs.get("text_color", self.button.cget("text_color")))
        )
        if self._text_color is None:
            self._text_color = "#ffffff"

        self._base_fg_color = self._normalize_color(
            self._resolve_single_color(kwargs.get("fg_color", self.button.cget("fg_color")))
        )
        self._hover_fg_color = self._darken_color(self._base_fg_color, 0.25)
        self.button.configure(fg_color=self._base_fg_color, hover_color=self._hover_fg_color)

        self.text_canvas = tk.Canvas(
            self.button,
            width=width,
            height=height,
            highlightthickness=0,
            bd=0,
            bg=self._base_fg_color,
        )
        self.text_canvas.place(relx=0.5, rely=0.5, anchor="center", relwidth=1, relheight=1)
        self.text_canvas.configure(cursor="hand2")
        self.text_canvas.bind("<Configure>", lambda event: self._draw_glow_text())
        self.text_canvas.bind("<ButtonRelease-1>", lambda _event: self.button.invoke())
        self.text_canvas.bind("<Enter>", self._on_enter)
        self.text_canvas.bind("<Leave>", self._on_leave)

        self._current_background = self._base_fg_color
        self._draw_glow_text()

        self.button.bind("<Enter>", self._on_enter)
        self.button.bind("<Leave>", self._on_leave)

    def _animate(self) -> None:
        if not self._animating:
            return
        self._offset = (self._offset + 2) % self._span
        self.canvas.coords(self.image_id, -self._offset, 0)
        self.after(50, self._animate)

    def _on_enter(self, _event=None) -> None:
        if not self._animating:
            self._animating = True
            self.canvas.itemconfigure(self.image_id, state="normal")
            self._animate()
        if self._current_background != self._hover_fg_color:
            self._current_background = self._hover_fg_color
            self.text_canvas.configure(bg=self._hover_fg_color)
            self._draw_glow_text()

    def _on_leave(self, _event=None) -> None:
        self._animating = False
        self.canvas.itemconfigure(self.image_id, state="hidden")
        self.canvas.coords(self.image_id, 0, 0)
        if self._current_background != self._base_fg_color:
            self._current_background = self._base_fg_color
            self.text_canvas.configure(bg=self._base_fg_color)
            self._draw_glow_text()

    def _draw_glow_text(self) -> None:
        self.text_canvas.delete("all")
        width = self.text_canvas.winfo_width()
        height = self.text_canvas.winfo_height()
        if width <= 0 or height <= 0:
            return
        center_x = width / 2
        center_y = height / 2
        glow_layers = (
            (1, 0.75),
            (2, 0.5),
            (3, 0.25),
        )
        for radius, strength in glow_layers:
            glow_color = self._blend_colors("#ffffff", self._current_background, strength)
            for dx in range(-radius, radius + 1):
                for dy in range(-radius, radius + 1):
                    if max(abs(dx), abs(dy)) == radius:
                        self.text_canvas.create_text(
                            center_x + dx,
                            center_y + dy,
                            text=self._text,
                            fill=glow_color,
                            font=self._font,
                        )
        self.text_canvas.create_text(
            center_x,
            center_y,
            text=self._text,
            fill=self._text_color,
            font=self._font,
        )

    def _resolve_single_color(self, color):
        if color is None:
            return None
        if isinstance(color, tuple):
            appearance = ctk.get_appearance_mode()
            return color[1] if appearance == "Dark" else color[0]
        return color

    def _normalize_color(self, color):
        if color is None:
            return None
        r, g, b = self._color_to_rgb(color)
        return self._rgb_to_hex((r, g, b))

    def _color_to_rgb(self, color):
        r, g, b = self.winfo_rgb(color)
        return (r // 256, g // 256, b // 256)

    def _rgb_to_hex(self, rgb) -> str:
        return f"#{rgb[0]:02x}{rgb[1]:02x}{rgb[2]:02x}"

    def _darken_color(self, color: str, amount: float) -> str:
        rgb = self._color_to_rgb(color)
        darkened = tuple(max(0, int(channel * (1 - amount))) for channel in rgb)
        return self._rgb_to_hex(darkened)

    def _blend_colors(self, foreground: str, background: str, alpha: float) -> str:
        alpha = max(0.0, min(1.0, alpha))
        fr, fg, fb = self._color_to_rgb(foreground)
        br, bg, bb = self._color_to_rgb(background)
        blended = (
            int(br * (1 - alpha) + fr * alpha),
            int(bg * (1 - alpha) + fg * alpha),
            int(bb * (1 - alpha) + fb * alpha),
        )
        return self._rgb_to_hex(blended)
