# -*- coding: utf-8 -*-
from sidebar.core.compat import tk, messagebox
from sidebar.core.config import VERSION

class FeedbackDialog(tk.Toplevel):
    def __init__(self, parent, outlook_client):
        tk.Toplevel.__init__(self, parent)
        self.outlook_client = outlook_client
        self.title("Send Feedback")
        # self.overrideredirect(True) # Removed to allow standard OS handling
        self.geometry("400x320")
        
        # Make Modal
        self.transient(parent)
        self.lift()
        self.attributes("-topmost", True)
        self.grab_set()

        # Center
        try:
             px = parent.winfo_rootx() + (parent.winfo_width() // 2) - 200
             py = parent.winfo_rooty() + (parent.winfo_height() // 2) - 160
             self.geometry("+{}+{}".format(px, py))
        except: pass

        # Dark Theme Colors
        self.colors = {
            "bg_root": "#202020",
            "bg_card": "#2D2D30",
            "accent": "#60CDFF",
            "fg_text": "#FFFFFF",
            "fg_dim": "#AAAAAA"
        }
        self.configure(bg=self.colors["bg_root"], highlightthickness=1, highlightbackground=self.colors["accent"])
        
        # Content
        content = tk.Frame(self, bg=self.colors["bg_root"], padx=15, pady=10)
        content.pack(fill="both", expand=True)
        
        tk.Label(content, text="Please describe your feedback, bug report, or suggestion:", 
                 fg=self.colors["fg_dim"], bg=self.colors["bg_root"], font=("Segoe UI", 9), anchor="w").pack(fill="x", pady=(0, 5))
        
        self.txt_input = tk.Text(content, bg=self.colors["bg_card"], fg=self.colors["fg_text"], 
                                 insertbackground="white", font=("Segoe UI", 10), height=8, relief="flat", padx=5, pady=5)
        self.txt_input.pack(fill="both", expand=True, pady=5)
        self.txt_input.focus_set()

        # Buttons
        btn_frame = tk.Frame(content, bg=self.colors["bg_root"])
        btn_frame.pack(fill="x", pady=10)
        
        def on_send():
            text = self.txt_input.get("1.0", "end-1c").strip()
            if not text:
                messagebox.showwarning("Empty", "Please enter some feedback first.")
                return
            
            # Send
            # Recipient: Ralph (Developer)
            recipient = "ralph.marsh@coveya.co.uk"
            subject = "Outlook Sidebar Feedback ({})".format(VERSION)
            
            success = self.outlook_client.send_email(subject, text, recipient)
            if success:
                messagebox.showinfo("Sent", "Feedback sent successfully!")
                self.destroy()
            else:
                messagebox.showerror("Error", "Failed to send feedback.\nPlease check Outlook is running.")

        btn_cancel = tk.Label(btn_frame, text="Cancel", fg="#AAAAAA", bg=self.colors["bg_root"], 
                              font=("Segoe UI", 10), cursor="hand2", padx=15, pady=5)
        btn_cancel.pack(side="right", padx=5)
        btn_cancel.bind("<Button-1>", lambda e: self.destroy())
        
        btn_send = tk.Label(btn_frame, text="Send Feedback", fg="white", bg=self.colors["accent"], 
                            font=("Segoe UI", 10, "bold"), cursor="hand2", padx=15, pady=5)
        btn_send.pack(side="right", padx=5)
        btn_send.bind("<Button-1>", lambda e: on_send())
        
        # Hover effects
        def on_enter(w, bg): w.config(bg=bg)
        def on_leave(w, bg): w.config(bg=bg)
        
        btn_send.bind("<Enter>", lambda e: on_enter(btn_send, "#40b0ff"))
        btn_send.bind("<Leave>", lambda e: on_leave(btn_send, self.colors["accent"]))
        
        btn_cancel.bind("<Enter>", lambda e: btn_cancel.config(fg="white"))
        btn_cancel.bind("<Leave>", lambda e: btn_cancel.config(fg="#AAAAAA"))
