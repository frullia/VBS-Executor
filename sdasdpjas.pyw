import tkinter as tk
import subprocess
import tempfile
import os

BG_COLOR = "#1e1e1e"
FG_COLOR = "#d4d4d4"
FONT = ("Segoe UI", 10)

def run_vbs_code():
    vbs_code = input_text.get("1.0", tk.END).strip()

    if not vbs_code:
        return

    with tempfile.NamedTemporaryFile(delete=False, suffix=".vbs", mode='w', encoding='utf-8') as temp_vbs:
        temp_vbs.write(vbs_code)
        temp_filename = temp_vbs.name

    try:
        
        subprocess.run(["cscript", "//nologo", temp_filename])
    except Exception:
        pass
    finally:
        os.remove(temp_filename)

def make_window_draggable(win):
    def start_move(event):
        win.x = event.x
        win.y = event.y

    def do_move(event):
        x = event.x_root - win.x
        y = event.y_root - win.y
        win.geometry(f"+{x}+{y}")

    win.bind("<ButtonPress-1>", start_move)
    win.bind("<B1-Motion>", do_move)

root = tk.Tk()
root.overrideredirect(True)
root.configure(bg=BG_COLOR)
root.geometry("400x220")
root.resizable(False, False)

make_window_draggable(root)

input_label = tk.Label(root, text="Enter VBScript code:", bg=BG_COLOR, fg=FG_COLOR, font=FONT)
input_label.pack(anchor="w", padx=10, pady=(10, 0))

input_text = tk.Text(root, height=8, width=50, bg=BG_COLOR, fg=FG_COLOR, insertbackground=FG_COLOR, font=FONT, relief=tk.FLAT, borderwidth=0)
input_text.pack(padx=10, pady=10)

run_button = tk.Button(root, text="Execute", command=run_vbs_code, bg="#0e639c", fg="white",
                       font=FONT, relief=tk.FLAT, activebackground="#1177cc", padx=15, pady=5, borderwidth=0)
run_button.pack(pady=(0, 10))

root.mainloop()
