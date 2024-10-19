import tkinter as tk
from pathlib import Path
from tkinter import filedialog
import win32com.client as win32
import csv
import json
import pydantic
from typing import List


ol = win32.Dispatch("Outlook.Application")  # Stores outlook connection
JSON_LOCATION = Path("assets/emails.json") # Stores email addresses in a json file


class Details(pydantic.BaseModel):
    name: str
    email: str
    subject_specific: str
    location: str


# ---------------------------------------------------------------------
def get_file_name() -> str:
    """
    Collects the name of the csv file required.
    :return: a path.
    """
    file = tk.Tk()
    file.withdraw()
    path:str = filedialog.askopenfilename(
        filetypes=(("csv files", "*.csv"), ("All files", "*.*"))
    )
    file.destroy()
    return path


# ---------------------------------------------------------------------
def set_text(widget, text: str) -> None:
    """Sets the value for text field
    :param widget: tkinter widget to change value
    :param text: new text to insert in tkinter widget
    :return: None
    """
    widget.delete(0, "end")
    widget.insert(0, text)


# ---------------------------------------------------------------------
def make_json(csv_location: str, json_location: Path) -> str|None:
    """Creates a json file from data provided
    :param csv_location: Path to csv file
    :param json_location: Path to json file
    :return: None
    """
    data: list = []
    try:
        with open(csv_location) as csvf:
            csv_reader = csv.DictReader(csvf)
            for rows in csv_reader:
                data.append(rows)
    except FileExistsError:
        return "Selected csv file does not exist"

    with open(json_location, "w", encoding="utf-8") as jsonf:
        jsonf.write(json.dumps(data, indent=4))


# ---------------------------------------------------------------------
def get_values(
    csv_path,
    cc_email,
    subject,
    message,
    signature,
) -> dict:
    """Gets values inputed in form by user
    :param csv_path: tkinter widget containing path for csv file.
    :param cc_email: tkinter widget containing email addressed to be cc'd.
    :param subject: tkinter widget containing subject of the emal.
    :param message: tkinter widget containing the email body.
    :param signature: tkinter widget containing the personalised email signature.
    :return: Python cictionary
    """

    # Create dict
    values: dict = {
        "to": csv_path.get(),
        "cc": cc_email.get(),
        "subject": subject.get(),
        "message": message.get("1.0", "end"),
        "signature": signature.get("1.0", "end"),
    }
    return values


# ---------------------------------------------------------------------
def send_email(user_inputs: dict, data: Details):
    """Sends emails to specified email addresses
    :param user_inputs: All user inputs
    :param data: All details captured from csv file
    :returns: None
    """
    olmailitem = 0 * 0  # store size of new mail
    mail = ol.CreateItem(olmailitem)  # new email reference
    mail.Subject = f"{user_inputs['subject']} | {data.subject_specific}"
    mail.To = data.email
    mail.CC = user_inputs["cc"]

    mail.HTMLBody = f'<p style="font-family:Calibri">Hi {data.name},</p>\
                    {user_inputs["message"]}\
                    <br>\
                    <br>\
                    {user_inputs["signature"]}'

    mail.Attachments.Add(data.location)
    return mail


def format_string(text: str) -> str:
    """
    Parses and returns text formated as a html element
    :param text: text to be parsed
    :return: a string
    """
    new_text: str = ""
    for line in text:
        new_text += (
            f'<p style="font-family:Calibri; margin: 0; padding: 0;"> {line} </p>'
        )
    return new_text


# ---------------------------------------------------------------------
def main(user_inputs: dict, preview: bool = False) -> None:
    """
    Main Function
    :param user_inputs: All input data from user
    :param preview: Boolean. if test email
    :return: None
    """
    # Create Json File
    make_json(user_inputs["to"], JSON_LOCATION)

    # Format message text
    user_inputs["message"] = format_string(user_inputs["message"].split("\n"))
    # Fornat signature text
    user_inputs["signature"] = format_string(user_inputs["signature"].split("\n"))

    with open(JSON_LOCATION) as file:
        json_data = json.load(file)
        data: List[Details] = [Details(**item) for item in json_data]

        if preview:  # if test only preview email
            mail = send_email(user_inputs=user_inputs, data=data[0])
            mail.Display()  # Preview email

        else:  # Send Email
            for line in data:
                mail = send_email(user_inputs=user_inputs, data=line)
                mail.Send()

    # clean up files created
    if JSON_LOCATION.is_file():
        JSON_LOCATION.unlink()


# ---------------------------------------------------------------------
if __name__ == "__main__":
    raise NotImplementedError()
