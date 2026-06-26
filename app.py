import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import requests

API_URL = "http://127.0.0.1:8000/api/transform"


def process_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])

    if not file_path:
        return

    df = pd.read_excel(file_path)

    # Replace NaN/inf with None so json serialization produces null (not NaN)
    df = df.where(pd.notnull(df), other=None)
    rows = df.to_dict(orient="records")

    try:
        response = requests.post(API_URL, json={"rows": rows})
        data = response.json()["data"]

        output_df = pd.DataFrame(data)

        save_path = filedialog.asksaveasfilename(defaultextension=".xlsx")

        if save_path:
            output_df.to_excel(save_path, index=False)
            messagebox.showinfo("Success", "File processed successfully!")

    except Exception as e:
        messagebox.showerror("Error", str(e))


root = tk.Tk()
root.title("Excel Transformer")

btn = tk.Button(root, text="Select Excel File", command=process_file)
btn.pack(pady=20)

root.mainloop()
