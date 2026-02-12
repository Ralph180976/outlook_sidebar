# -*- coding: utf-8 -*-
from sidebar.core.compat import tk, ttk

class ScrollableFrame(tk.Frame):
    """
    A scrollable frame that can contain multiple email cards.
    """
    def __init__(self, container, *args, auto_hide_scrollbar=True, **kwargs):
        tk.Frame.__init__(self, container, *args, **kwargs)
        self._auto_hide = auto_hide_scrollbar
        self.canvas = tk.Canvas(self, bg=kwargs.get("bg", "#222222"), highlightthickness=0)
        self.scrollable_frame = tk.Frame(self.canvas, bg=kwargs.get("bg", "#222222"))

        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(
                scrollregion=self.canvas.bbox("all")
            )
        )
        
        # Speed up scrolling (Windows default is slow)
        self.canvas.configure(yscrollincrement=20)

        self.window_id = self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")

        # Scrollbar on the right side
        self.scrollbar = tk.Scrollbar(self, orient="vertical", command=self.canvas.yview,
                                      bg="#444444", troughcolor="#2D2D30", width=8,
                                      highlightthickness=0, bd=0)

        # Pack scrollbar FIRST to reserve its space, then canvas fills the rest
        self.scrollbar.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)

        if self._auto_hide:
            self.canvas.configure(yscrollcommand=self._on_scroll_update)
            self._scrollbar_visible = True  # starts visible, will hide if not needed
        else:
            self.canvas.configure(yscrollcommand=self.scrollbar.set)
            self._scrollbar_visible = True
        
        # Mousewheel scrolling
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)
        
        # Ensure scrollable frame matches canvas width
        self.canvas.bind("<Configure>", self._on_canvas_configure)

    def _on_scroll_update(self, first, last):
        """Show scrollbar only when content overflows the visible area."""
        self.scrollbar.set(first, last)
        try:
            f, l = float(first), float(last)
            needs_scroll = not (f <= 0.001 and l >= 0.999)
            if needs_scroll and not self._scrollbar_visible:
                self.scrollbar.pack(side="right", fill="y")
                self.canvas.pack_forget()
                self.canvas.pack(side="left", fill="both", expand=True)
                self._scrollbar_visible = True
            elif not needs_scroll and self._scrollbar_visible:
                self.scrollbar.pack_forget()
                self._scrollbar_visible = False
        except:
            pass

    def _on_mousewheel(self, event):
        self.canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        
    def _on_canvas_configure(self, event):
        # Resize the inner frame to match the canvas width
        self.canvas.itemconfig(self.window_id, width=event.width)

    def config(self, **kwargs):
        """Propagate configuration to internal components."""
        bg = kwargs.get("bg") or kwargs.get("background")
        if bg:
            self.canvas.config(bg=bg)
            self.scrollable_frame.config(bg=bg)
        # Apply to self (the container frame)
        tk.Frame.config(self, **kwargs)

    def configure(self, **kwargs):
        self.config(**kwargs)


class RoundedFrame(tk.Canvas):
    def __init__(self, parent, width, height, corner_radius, padding, color, bg, **kwargs):
        tk.Canvas.__init__(self, parent, width=width, height=height, bg=bg, bd=0, highlightthickness=0, **kwargs)
        self.radius = corner_radius
        self.padding = padding
        self.color = color
        
        self.id = self.create_rounded_rect(0, 0, width, height, self.radius, fill=self.color, outline="")
        
        # Inner frame for widgets
        self.inner = tk.Frame(self, bg=self.color)
        self.window_id = self.create_window((padding, padding), window=self.inner, anchor="nw")
        
        self.bind("<Configure>", self._on_resize)
        
    def _on_resize(self, event):
        self.coords(self.id, self._rounded_rect_coords(0, 0, event.width, event.height, self.radius))
        self.itemconfig(self.window_id, width=event.width - 2*self.padding, height=event.height - 2*self.padding)
        
    def create_rounded_rect(self, x1, y1, x2, y2, r, **kwargs):
        return self.create_polygon(self._rounded_rect_coords(x1, y1, x2, y2, r), **kwargs)

    def _rounded_rect_coords(self, x1, y1, x2, y2, r):
        points = [x1+r, y1,
                  x1+r, y1,
                  x2-r, y1,
                  x2-r, y1,
                  x2, y1,
                  x2, y1+r,
                  x2, y1+r,
                  x2, y2-r,
                  x2, y2-r,
                  x2, y2,
                  x2-r, y2,
                  x2-r, y2,
                  x1+r, y2,
                  x1+r, y2,
                  x1, y2,
                  x1, y2-r,
                  x1, y2-r,
                  x1, y1+r,
                  x1, y1+r,
                  x1, y1]
        return points

class ToolTip:
    """
    Creates a popup tooltip for a given widget.
    """
    def __init__(self, widget, text, side="bottom"):
        self.widget = widget
        self.text = text
        self.side = side # "bottom", "left", "right", "top"
        self.tip_window = None
        self.widget.bind("<Enter>", self.enter)
        self.widget.bind("<Leave>", self.leave)

    def enter(self, event=None):
        self.show_tip()

    def leave(self, event=None):
        self.hide_tip()

    def show_tip(self):
        """Displays the tooltip."""
        if self.tip_window or not self.text:
            return
        
        # Create window first to get size
        self.tip_window = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True)
        tw.wm_attributes("-topmost", True)
        
        label = tk.Label(
            tw, 
            text=self.text, 
            justify="left",
            bg="#2d2d2d", 
            fg="#ffffff",
            relief="solid", 
            borderwidth=1,
            font=("Segoe UI", 8),
            padx=4, pady=2
        )
        label.pack(ipadx=1)
        
        tw.update_idletasks() # Calculate size
        
        tw_width = tw.winfo_reqwidth()
        tw_height = tw.winfo_reqheight()
        
        widget_x = self.widget.winfo_rootx()
        widget_y = self.widget.winfo_rooty()
        widget_w = self.widget.winfo_width()
        widget_h = self.widget.winfo_height()
        
        if self.side == "left":
            x = widget_x - tw_width - 5
            y = widget_y + (widget_h // 2) - (tw_height // 2)
        elif self.side == "right":
            x = widget_x + widget_w + 5
            y = widget_y + (widget_h // 2) - (tw_height // 2)
        elif self.side == "top":
            x = widget_x + (widget_w // 2) - (tw_width // 2)
            y = widget_y - tw_height - 5
        else: # bottom
            x = widget_x + 20
            y = widget_y + widget_h + 5
            
        tw.wm_geometry("+{}+{}".format(x, y))

    def hide_tip(self):
        """Hides the tooltip."""
        if self.tip_window:
            self.tip_window.destroy()
            self.tip_window = None
