from pathlib import Path

# from tkinter import *
# Explicit imports to satisfy Flake8
from tkinter import Tk, Canvas, Entry, Text, Button, PhotoImage, filedialog, messagebox
import pandas as pd
from pymavlink import mavutil # type: ignore

OUTPUT_PATH = Path(__file__).parent
ASSETS_PATH = OUTPUT_PATH / Path(r"/Users/vishalmehra/Desktop/build/assets/frame0")


def relative_to_assets(path: str) -> Path:
    return ASSETS_PATH / Path(path)

def parse_ardupilot_log(file_path):
    # Initialize a connection to the log file
    mlog = mavutil.mavlink_connection(file_path)

    # Dictionary to hold all parsed data
    data = {}

    # Read messages from the log
    while True:
        try:
            msg = mlog.recv_match(blocking=False)
            if msg is None:
                break

            # Skip heartbeat messages to avoid excessive data
            if msg.get_type() == 'HEARTBEAT':
                continue

            # Add message to data dictionary
            msg_dict = msg.to_dict()
            msg_type = msg.get_type()

            if msg_type not in data:
                data[msg_type] = []

            data[msg_type].append(msg_dict)
        except Exception as e:
            print(f"Error: {e}")
            break

    return data

def convert_to_excel(data, output_file):
    # Create a Pandas Excel writer using openpyxl as the engine
    writer = pd.ExcelWriter(output_file, engine='openpyxl')

    # Convert each message type to a separate sheet in the Excel file
    for msg_type, msg_list in data.items():
        df = pd.DataFrame(msg_list)
        df.to_excel(writer, sheet_name=msg_type, index=False)

    # Save the Excel file
    writer.close()

def select_file():
    # Open file dialog to select the .bin file
    file_path = filedialog.askopenfilename(filetypes=[("BIN files", "*.bin")])
    if file_path:
        output_file = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if output_file:
            try:
                data = parse_ardupilot_log(file_path)
                convert_to_excel(data, output_file)
                messagebox.showinfo("Success", f"Data successfully written to {output_file}")
            except Exception as e:
                messagebox.showerror("Error", f"An error occurred: {e}")


window = Tk()

window.geometry("400x300")
window.configure(bg = "#FFFFFF")


canvas = Canvas(
    window,
    bg = "#FFFFFF",
    height = 300,
    width = 400,
    bd = 0,
    highlightthickness = 0,
    relief = "ridge"
)

canvas.place(x = 0, y = 0)
image_image_1 = PhotoImage(
    file=relative_to_assets("image_1.png"))
image_1 = canvas.create_image(
    200.0,
    150.0,
    image=image_image_1
)

button_image_1 = PhotoImage(
    file=relative_to_assets("button_1.png"))
button_1 = Button(
    image=button_image_1,
    borderwidth=0,
    highlightthickness=0,
    command=select_file,
    relief="flat"
)
button_1.place(
    x=116.0,
    y=208.0,
    width=169.0,
    height=52.0
)

image_image_2 = PhotoImage(
    file=relative_to_assets("image_2.png"))
image_2 = canvas.create_image(
    200.0,
    168.0,
    image=image_image_2
)

image_image_3 = PhotoImage(
    file=relative_to_assets("image_3.png"))
image_3 = canvas.create_image(
    200.0,
    104.0,
    image=image_image_3
)

image_image_4 = PhotoImage(
    file=relative_to_assets("image_4.png"))
image_4 = canvas.create_image(
    200.0,
    47.0,
    image=image_image_4
)
window.resizable(False, False)
window.mainloop()
