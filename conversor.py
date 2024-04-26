import tkinter as tk
from tkinter import filedialog, messagebox
from icalendar import Calendar
import pandas as pd
from datetime import datetime

def convertir_y_ordenar():
    file_path = filedialog.askopenfilename(filetypes=[("iCalendar files", "*.ics")])
    if not file_path: 
        return

    with open(file_path, 'rb') as ics_file:
        gcal = Calendar.from_ical(ics_file.read())

    events_data = []
    for component in gcal.walk():
        if component.name == "VEVENT":
            summary = component.get('summary', '')
            start = component.get('dtstart').dt
            end = component.get('dtend').dt
            description = component.get('description', '')
            start = make_naive(start)
            end = make_naive(end)
            category = categorize_event(summary)
            events_data.append([summary, start, end, description, category])

    df = pd.DataFrame(events_data, columns=["Summary", "Start", "End", "Description", "Category"])
    df['Start'] = pd.to_datetime(df['Start'], errors='coerce')
    df['End'] = pd.to_datetime(df['End'], errors='coerce')
    df['Year'] = df['Start'].dt.year
    df['Month'] = df['Start'].dt.month

    output_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    if not output_path:
        return

    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        for (year, month), group in df.groupby(['Year', 'Month']):
            sheet_name = f"{year}-{month:02d}"
            taxis = group[group['Category'] == 'taxis'].drop(columns='Category').reset_index(drop=True)
            tours = group[group['Category'] == 'tours'].drop(columns='Category').reset_index(drop=True)
            otros = group[group['Category'] == 'otros'].drop(columns='Category').reset_index(drop=True)
            month_data = pd.DataFrame({
                'Taxis Summary': taxis['Summary'],
                'Taxis Start': taxis['Start'],
                'Taxis End': taxis['End'],
                'Taxis Description': taxis['Description'],
                'Tours Summary': tours['Summary'],
                'Tours Start': tours['Start'],
                'Tours End': tours['End'],
                'Tours Description': tours['Description'],
                'Otros Summary': otros['Summary'],
                'Otros Start': otros['Start'],
                'Otros End': otros['End'],
                'Otros Description': otros['Description']
            })
            month_data.to_excel(writer, sheet_name=sheet_name, index=False)

    messagebox.showinfo("Éxito", "Archivo convertido y ordenado en XLSX con éxito.")

def make_naive(value):
    if isinstance(value, datetime) and value.tzinfo is not None and value.tzinfo.utcoffset(value) is not None:
        return value.replace(tzinfo=None)
    return value

def categorize_event(summary):
    if summary is None:
        return "otros"
    summary_lower = summary.lower()
    if "taxi" in summary_lower:
        return "taxis"
    elif "tour" in summary_lower:
        return "tours"
    else:
        return "otros"

root = tk.Tk()
root.title("")
root.geometry("800x600")

btn_convertir_ordenar = tk.Button(root, text="Convertir y Ordenar", command=convertir_y_ordenar)
btn_convertir_ordenar.pack(pady=20)

label_marca = tk.Label(root, text="SanjayGRR \u00AE", fg="white", bg="black")
label_marca.pack(side="bottom", fill="x")

root.mainloop()
