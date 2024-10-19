import ttkbootstrap as ttk
import tkinter as tk
from ttkbootstrap.constants import *
from util import get_file_name, set_text, main, get_values


# Label widget class
class LabelWidget(ttk.Label):
    def __init__(self, master=None, **kwargs) -> None:
        super().__init__(master, **kwargs)
        self.label = ttk.Label(master, text=kwargs.get("text"), width=8)
        self.label.pack(pady=10, side="left", padx=5, anchor="n")


# --------------------------------------------------------------------------------------


def draw_window(window) -> None:
    """
    Creates window for app
    :param window: tkinter window widget
    :return: None
    """
    # --------- Header Logo ---------- #
    logo_frame = ttk.Frame(window)
    logo_frame.pack(pady=5)
    logo_widget = tk.Label(logo_frame, image=ICON, width=400, height=100)
    logo_widget.pack()
    # Label
    logo_label = ttk.Label(logo_frame, text="Bulk Email Handler")
    logo_label.pack()
    logo_label.config(font=("Arial", 20, "bold"))

    # --------- Form Widgets -----------#
    # email csv list
    email_list_frame = ttk.Frame(window)
    email_list_frame.pack(pady=4, padx=10, fill="x")
    LabelWidget(email_list_frame, text="TO")
    email_list = ttk.Entry(email_list_frame)
    email_list.pack(side="left", fill="x", expand=True, padx=8)
    email_list.insert(0, "Enter path to csv file")

    # Button to open file browser
    btn_frame = ttk.Frame(window)
    btn_frame.pack(pady=1, padx=20, anchor="e")
    btn_browse = ttk.Button(
        btn_frame,
        text="Browse",
        bootstyle=SECONDARY,
        command=lambda: set_text(email_list, get_file_name()),
    )
    btn_browse.pack()

    # cc list
    cc_frame = ttk.Frame(window)
    cc_frame.pack(pady=6, padx=10, fill="x")
    LabelWidget(cc_frame, text="CC")
    cc_email = ttk.Entry(cc_frame)
    cc_email.pack(side="left", fill="x", expand=True, padx=8)

    # Subject
    subject_frame = ttk.Frame(window)
    subject_frame.pack(pady=6, padx=10, fill="x")
    LabelWidget(subject_frame, text="Subject")
    subject = ttk.Entry(subject_frame)
    subject.pack(side="left", fill="x", expand=True, padx=8)

    # Message
    message_frame = ttk.Frame(window)
    message_frame.pack(pady=6, padx=10, fill="x")
    LabelWidget(message_frame, text="Message")
    message = tk.Text(message_frame, height=7)
    message.pack(side="left", fill="x", padx=8)

    # Signature
    signaure_frame = ttk.Frame(window)
    signaure_frame.pack(pady=6, padx=10, fill="x")
    LabelWidget(signaure_frame, text="Signature")
    signaure = tk.Text(signaure_frame, height=7)
    signaure.pack(side="left", fill="x", padx=8)

    # ------------ Buttons ---------------- #
    button_frame = ttk.Frame(window)
    button_frame.pack(pady=6, padx=10, anchor="center")

    # Test button
    test_btn = ttk.Button(
        button_frame,
        text="Preview Email",
        bootstyle="success-outline",
        command=lambda: main(
            get_values(email_list, cc_email, subject, message, signaure), test=True
        ),
    )
    test_btn.pack(pady=10, side="left", padx=5)

    send_btn = ttk.Button(
        button_frame,
        text="Send Emails",
        bootstyle="primary",
        command=lambda: main(
            get_values(email_list, cc_email, subject, message, signaure)
        ),
    )
    send_btn.pack(pady=10, side="left", padx=5)
    root.mainloop()


# -----------------------------------------------------------------------------
def center_app(window, width: int, height: int) -> None:
    """
    Center app on screen
    :param window: tkinter window widget
    :param width: width of window
    :param height: height of window
    :return: None
    """
    screen_width = window.winfo_screenwidth()  # Get screen width
    screen_height = window.winfo_screenheight()  # egt screen height
    # calculate app position
    x = (screen_width // 2) - (width // 2)
    y = (screen_height // 2) - (height // 2)
    # Set window size plus position
    window.geometry(f"{width}x{height}+{x}+{y}")


# -------------------------------------------------------------------------
if __name__ == "__main__":

    # Constants
    APP_WIDTH: int = 660
    APP_HEIGHT: int = 755

    root = ttk.Window(themename="minty")
    root.geometry(f"{APP_WIDTH}x{APP_HEIGHT}")
    center_app(root, APP_WIDTH, APP_HEIGHT)
    root.minsize(650, 750)
    root.maxsize(660, 755)
    root.title("Bulk Email Handler")
    ICON = tk.PhotoImage(file="assets/logo.png")
    root.iconphoto(False, ICON)
    draw_window(root)
