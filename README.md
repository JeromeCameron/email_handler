# Email Handler

*I needed a way to send bulk emails with unique attachments, 
new a little python and this is the result.*

This Python app allows users to send bulk emails with unique attachments to 
multiple recipients. Each email has a custom attachment, making it ideal for 
personalized communications like reports, certificates, or other documents.
This uses the pywin32 library to send emails through your locally installed ms outlook.

## Installation

```bash
pip install -r requirements.txt
````

## Usage
```python
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
```

1. Prepare the email list: The names, email addresses and attachment location should be 
stored in a CSV.
2. Run the script:

```bash
python app.py
```

![Screeshot of the app!](/images/screenshot.jpg "App")


## Dependencies
- pillow==10.4.0
- pydantic==2.8.2
- pydantic_core==2.20.1
- pywin32==306
- tk==0.1.0
- ttkbootstrap==1.10.1
