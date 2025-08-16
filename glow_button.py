import tkinter as tk
import customtkinter as ctk


class GlowButton(ctk.CTkFrame):
    """Button with animated gradient border on hover."""

    def __init__(self, master, text: str, command=None, width: int = 140, height: int = 32, **kwargs) -> None:
        border = 2
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
            text=text,
            command=command,
            width=width,
            height=height,
            corner_radius=10,
            **kwargs,
        )
        self.button.place(x=border, y=border)

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

    def _on_leave(self, _event=None) -> None:
        self._animating = False
        self.canvas.itemconfigure(self.image_id, state="hidden")
        self.canvas.coords(self.image_id, 0, 0)
