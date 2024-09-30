import win32gui
import win32api
import win32con
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox

# UI Styling
root = tk.Tk()
root.title("Window Opacity Control")
root.configure(bg="#2b2b2b")  # Dark background
style = ttk.Style()
style.theme_use("clam")  # Use a more modern theme if available on the system.
style.configure("TButton", background="#4a4a4a", foreground="white", relief="flat", padding=6)
style.configure("TCheckbutton", background="#2b2b2b", foreground="white")
style.configure("TScale", troughcolor="#4a4a4a", background="#2b2b2b", foreground="white")
style.configure("TLabel", background="#2b2b2b", foreground="white")
style.configure("TFrame", background="#2b2b2b")
style.configure("TListbox", background="#383838", foreground="white", selectbackground="#555555")

def get_windows():
    """Retrieves a list of currently active window titles and handles."""
    windows = []

    def enum_windows_proc(hwnd, lparam):
        if win32gui.IsWindowVisible(hwnd) and win32gui.GetWindowText(hwnd):
            windows.append((win32gui.GetWindowText(hwnd), hwnd))
        return True

    win32gui.EnumWindows(enum_windows_proc, None)
    return windows

def set_opacity(hwnd, opacity, click_through=False):
    """Sets opacity and click-through, robustly handling style restoration."""
    original_style = original_styles.get(hwnd, 0)  # Get original or 0 if not stored

    if hwnd not in original_styles:
        original_styles[hwnd] = win32gui.GetWindowLong(hwnd, win32con.GWL_EXSTYLE)

    new_style = original_style | win32con.WS_EX_LAYERED  # Apply layered style

    if click_through:
        new_style |= win32con.WS_EX_TRANSPARENT  # Allow clicks to pass through
    else:
        new_style &= ~win32con.WS_EX_TRANSPARENT  # Disable click-through if unchecked

    win32gui.SetWindowLong(hwnd, win32con.GWL_EXSTYLE, new_style)

    # Set the opacity
    alpha = int(255 * opacity / 100)  # Correct alpha calculation
    win32gui.SetLayeredWindowAttributes(hwnd, 0, alpha, win32con.LWA_ALPHA)

def set_always_on_top(hwnd, always_on_top):
    """Sets the always-on-top attribute."""
    flag = win32con.HWND_TOPMOST if always_on_top else win32con.HWND_NOTOPMOST
    win32gui.SetWindowPos(hwnd, flag, 0, 0, 0, 0, win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)

def refresh_windows():
    """Refreshes the list of windows in the UI."""
    global windows  # Access the global windows variable
    window_list.delete(0, tk.END)
    windows = get_windows()  # Update the global windows list
    for title, hwnd in windows:
        window_list.insert(tk.END, title)

def apply_settings():
    try:
        selected_index = window_list.curselection()[0]
        title = window_list.get(selected_index)
        hwnd = next((hwnd for win_title, hwnd in windows if win_title == title), None)
        if not hwnd: raise IndexError

        opacity = opacity_slider.get()
        click_through = click_through_var.get()
        set_opacity(hwnd, opacity, click_through)

        always_on_top = always_on_top_var.get()
        set_always_on_top(hwnd, always_on_top)

    except IndexError:
        messagebox.showwarning("Error", "Select a window (or it may have closed).")


original_styles = {}

# --- UI Layout ---
main_frame = ttk.Frame(root, padding="20")
main_frame.pack(fill=tk.BOTH, expand=True)

window_list = tk.Listbox(main_frame, width=50, height=10, font=("TkDefaultFont", 10))  # Increased height
window_list.pack(pady=(0, 10), fill=tk.X)

opacity_label = ttk.Label(main_frame, text="Opacity:", font=("TkDefaultFont", 10, "bold"))  # Label
opacity_label.pack()
opacity_slider = ttk.Scale(main_frame, from_=10, to=100, orient=tk.HORIZONTAL, length=250)  # Longer
opacity_slider.set(100)
opacity_slider.pack(pady=(0, 10))

checkbox_frame = ttk.Frame(main_frame)
checkbox_frame.pack()

always_on_top_var = tk.BooleanVar()
always_on_top_check = ttk.Checkbutton(checkbox_frame, text="Always on Top", variable=always_on_top_var)
always_on_top_check.pack(side=tk.LEFT, padx=5)

click_through_var = tk.BooleanVar()
click_through_check = ttk.Checkbutton(checkbox_frame, text="Click-Through", variable=click_through_var)
click_through_check.pack(side=tk.LEFT, padx=5)

apply_button = ttk.Button(main_frame, text="Apply", command=apply_settings)  # ttk button
apply_button.pack(pady=(15, 5))

refresh_button = ttk.Button(main_frame, text="Refresh", command=refresh_windows)
refresh_button.pack()

windows = get_windows()
refresh_windows()

root.mainloop()
