import tkinter as tk

class ToolTip:
    """
    Creates a popup tooltip for a given widget.
    """
    def __init__(self, widget, text):
        self.widget = widget
        self.text = text
        self.tip_window = None
        self.schedule_id = None
        self.widget.bind("<Enter>", self.enter)
        self.widget.bind("<Leave>", self.leave)

    def enter(self, event=None):
        self.schedule_id = self.widget.after(1000, self.show_tip)

    def leave(self, event=None):
        if self.schedule_id:
            self.widget.after_cancel(self.schedule_id)
            self.schedule_id = None
        self.hide_tip()

    def show_tip(self):
        """Displays the tooltip."""
        if self.tip_window or not self.text:
            return
        
        x, y, _, _ = self.widget.bbox("insert")
        x = x + self.widget.winfo_rootx() + 20
        y = y + self.widget.winfo_rooty() + 20
        
        self.tip_window = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True) # Removes window decorations
        tw.wm_geometry(f"+{x}+{y}")
        
        label = tk.Label(
            tw, 
            text=self.text, 
            justify="left",
            bg="#ffffe0", 
            fg="#000000",
            relief="solid", 
            borderwidth=1,
            font=("Segoe UI", 9, "normal"),
            padx=4, pady=2
        )
        label.pack(ipadx=1)

    def hide_tip(self):
        """Hides the tooltip."""
        if self.tip_window:
            self.tip_window.destroy()
            self.tip_window = None
