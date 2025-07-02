"""Small animation helpers for the customtkinter GUI."""

import customtkinter as ctk
from tkinter import messagebox


def fade_window(root: ctk.CTk, callback, steps: int = 10, delay: int = 30) -> None:
    """Fade the entire window out, run callback, then fade back in."""
    def fade_out(step: int = 0) -> None:
        if step <= steps:
            root.attributes("-alpha", 1 - step / steps)
            root.after(delay, fade_out, step + 1)
        else:
            callback()
            fade_in()

    def fade_in(step: int = 0) -> None:
        if step <= steps:
            root.attributes("-alpha", step / steps)
            root.after(delay, fade_in, step + 1)
        else:
            root.attributes("-alpha", 1.0)

    fade_out()


def pulse(widget: ctk.CTkBaseClass, color: str = "#90C2FF", duration: int = 200) -> None:
    """Temporarily change the widget color to give feedback."""
    orig = widget.cget("fg_color")
    widget.configure(fg_color=color)
    widget.after(duration, lambda: widget.configure(fg_color=orig))


def fade_log(textbox: ctk.CTkTextbox, tag: str, steps: int = 10, delay: int = 50) -> None:
    """Fade a tagged log line from gray to black to draw attention."""
    def step(i: int) -> None:
        level = 55 + int((200 / steps) * i)
        color = f"#{level:02x}{level:02x}{level:02x}"
        textbox.tag_config(tag, foreground=color)
        if i < steps:
            textbox.after(delay, step, i + 1)
    step(0)


def slide_window(win: ctk.CTkToplevel, end_y: int, step: int = 10, delay: int = 10) -> None:
    """Animate a toplevel window sliding down from the top."""
    x = win.winfo_x()
    start = -win.winfo_height()
    pos = start

    def anim() -> None:
        nonlocal pos
        if pos < end_y:
            win.geometry(f"+{x}+{pos}")
            pos += step
            win.after(delay, anim)
        else:
            win.geometry(f"+{x}+{end_y}")
    anim()


def popup(root: ctk.CTk, title: str, message: str) -> None:
    """Show an informational popup with a slight slide animation."""
    win = ctk.CTkToplevel(root)
    win.title(title)
    ctk.CTkLabel(win, text=message, justify="left", wraplength=300).pack(padx=20, pady=20)
    win.update_idletasks()
    x = root.winfo_x() + (root.winfo_width() // 2 - win.winfo_width() // 2)
    slide_window(win, root.winfo_y() + 100)
    win.geometry(f"+{x}+{-win.winfo_height()}")
    win.after(2500, win.destroy)
