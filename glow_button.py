import math
import tkinter as tk
import customtkinter as ctk


class GlowButton(ctk.CTkFrame):
    """Button with animated gradient border on hover."""

    def __init__(
        self,
        master,
        text: str,
        command=None,
        width: int = 140,
        height: int = 32,
        **kwargs,
    ) -> None:
        border = float(kwargs.pop("border_thickness", 0.4))
        gradient_colors = kwargs.pop("gradient_colors", ("#ff0080", "#ff00c8"))
        glow_color = kwargs.pop("glow_color", "#ff3cff")
        glow_layers = kwargs.pop("glow_layers", ((2, 0.65), (4, 0.45), (6, 0.25)))
        hover_glow_multiplier = float(kwargs.pop("hover_glow_multiplier", 1.35))
        idle_glow_multiplier = float(kwargs.pop("idle_glow_multiplier", 1.0))
        animation_speed = int(kwargs.pop("animation_speed", 2))
        corner_radius = kwargs.pop("corner_radius", 10)

        total_width = width + border * 2
        total_height = height + border * 2

        super().__init__(master, width=total_width, height=total_height, fg_color="transparent")
        self._border = max(0.0, border)
        self._button_width = width
        self._button_height = height
        self._span = total_width
        self._span_px = max(1, int(math.ceil(total_width)))
        self._image_height = max(1, int(math.ceil(total_height)))
        self._gradient_width = self._span_px * 2
        self._offset = 0
        self._animating = False
        self._gradient_colors = gradient_colors
        self._glow_color = glow_color
        self._glow_layers = glow_layers
        self._hover_glow_multiplier = hover_glow_multiplier
        self._idle_glow_multiplier = idle_glow_multiplier
        self._current_glow_multiplier = idle_glow_multiplier
        self._animation_speed = max(1, animation_speed)
        self._corner_radius = corner_radius
        self._border_mask_id = None

        # Canvas for gradient border
        self.canvas = tk.Canvas(
            self,
            width=total_width,
            height=total_height,
            highlightthickness=0,
            bd=0,
        )
        self.canvas.place(x=0, y=0)

        # Create gradient image
        normalized_colors = tuple(self._normalize_color(color) for color in self._gradient_colors)
        start_rgb = self._color_to_rgb(normalized_colors[0])
        end_rgb = self._color_to_rgb(normalized_colors[-1])
        gradient_range = max(1, self._span_px - 1)
        self.gradient_image = tk.PhotoImage(width=self._gradient_width, height=self._image_height)
        for x in range(self._gradient_width):
            ratio = (x % self._span_px) / gradient_range
            color_rgb = (
                int(start_rgb[0] + (end_rgb[0] - start_rgb[0]) * ratio),
                int(start_rgb[1] + (end_rgb[1] - start_rgb[1]) * ratio),
                int(start_rgb[2] + (end_rgb[2] - start_rgb[2]) * ratio),
            )
            color_hex = self._rgb_to_hex(color_rgb)
            for y in range(self._image_height):
                self.gradient_image.put(color_hex, (x, y))
        self.image_id = self.canvas.create_image(0, 0, image=self.gradient_image, anchor="nw")
        self.canvas.itemconfigure(self.image_id, state="hidden")
        self.canvas.tag_lower(self.image_id)

        if "border_width" not in kwargs:
            kwargs["border_width"] = 0

        # Underlying CTkButton
        self.button = ctk.CTkButton(
            self,
            text="",
            command=command,
            width=width,
            height=height,
            corner_radius=self._corner_radius,
            **kwargs,
        )
        self.button.place(x=self._border, y=self._border)

        self._text = text
        self._font = kwargs.get("font", self.button.cget("font"))
        self._text_color = self._normalize_color(
            self._resolve_single_color(kwargs.get("text_color", self.button.cget("text_color")))
        )
        if self._text_color is None:
            self._text_color = "#ffffff"

        self._glow_color = self._normalize_color(self._glow_color)
        self._glow_layers = tuple(
            (max(0, int(round(radius))), max(0.0, float(strength)))
            for radius, strength in self._glow_layers
        )

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
        self._create_border_mask(self._base_fg_color)
        self._draw_glow_text()

        self.button.bind("<Enter>", self._on_enter)
        self.button.bind("<Leave>", self._on_leave)

    def _animate(self) -> None:
        if not self._animating:
            return
        self._offset = (self._offset + self._animation_speed) % self._span_px
        self.canvas.coords(self.image_id, -self._offset, 0)
        self.after(50, self._animate)

    def _on_enter(self, _event=None) -> None:
        if not self._animating:
            self._animating = True
            self.canvas.itemconfigure(self.image_id, state="normal")
            if self._border_mask_id is not None:
                self.canvas.tag_raise(self._border_mask_id)
            self._animate()
        redraw = False
        if self._current_background != self._hover_fg_color:
            self._current_background = self._hover_fg_color
            self.text_canvas.configure(bg=self._hover_fg_color)
            self._update_border_mask_fill(self._hover_fg_color)
            redraw = True
        if self._current_glow_multiplier != self._hover_glow_multiplier:
            self._current_glow_multiplier = self._hover_glow_multiplier
            redraw = True
        if redraw:
            self._draw_glow_text()

    def _on_leave(self, _event=None) -> None:
        self._animating = False
        self.canvas.itemconfigure(self.image_id, state="hidden")
        self.canvas.coords(self.image_id, 0, 0)
        if self._border_mask_id is not None:
            self.canvas.tag_raise(self._border_mask_id)
        redraw = False
        if self._current_background != self._base_fg_color:
            self._current_background = self._base_fg_color
            self.text_canvas.configure(bg=self._base_fg_color)
            self._update_border_mask_fill(self._base_fg_color)
            redraw = True
        if self._current_glow_multiplier != self._idle_glow_multiplier:
            self._current_glow_multiplier = self._idle_glow_multiplier
            redraw = True
        if redraw:
            self._draw_glow_text()

    def _draw_glow_text(self) -> None:
        self.text_canvas.delete("all")
        width = self.text_canvas.winfo_width()
        height = self.text_canvas.winfo_height()
        if width <= 0 or height <= 0:
            return
        center_x = width / 2
        center_y = height / 2
        for radius, strength in self._glow_layers:
            if radius <= 0 or strength <= 0:
                continue
            glow_strength = min(1.0, strength * self._current_glow_multiplier)
            glow_color = self._blend_colors(self._glow_color, self._current_background, glow_strength)
            for dx in range(-radius, radius + 1):
                for dy in range(-radius, radius + 1):
                    if dx * dx + dy * dy <= radius * radius:
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

    def _create_border_mask(self, fill_color: str) -> None:
        if self._border <= 0:
            return
        if self._border_mask_id is not None:
            self.canvas.delete(self._border_mask_id)
            self._border_mask_id = None
        self._border_mask_id = self._draw_rounded_rect(
            self._border,
            self._border,
            self._border + self._button_width,
            self._border + self._button_height,
            self._corner_radius,
            fill=fill_color,
            outline="",
        )
        self.canvas.tag_raise(self._border_mask_id)

    def _update_border_mask_fill(self, fill_color: str) -> None:
        if self._border_mask_id is not None:
            self.canvas.itemconfigure(self._border_mask_id, fill=fill_color)

    def _draw_rounded_rect(self, x1, y1, x2, y2, radius, **kwargs):
        radius = max(0.0, float(radius))
        if radius <= 0:
            return self.canvas.create_rectangle(x1, y1, x2, y2, **kwargs)

        radius = min(radius, (x2 - x1) / 2, (y2 - y1) / 2)
        points = [
            x1 + radius,
            y1,
            x2 - radius,
            y1,
            x2,
            y1,
            x2,
            y1 + radius,
            x2,
            y2 - radius,
            x2,
            y2,
            x2 - radius,
            y2,
            x1 + radius,
            y2,
            x1,
            y2,
            x1,
            y2 - radius,
            x1,
            y1 + radius,
            x1,
            y1,
        ]
        return self.canvas.create_polygon(points, smooth=True, splinesteps=36, **kwargs)

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
