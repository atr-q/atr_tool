import duckdb
import pandas
import math
import os
import sys
import numpy
import subprocess as sp
import docx
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.shared import Pt
from docx.shared import Inches, Cm
import customtkinter
from PIL import Image, ImageTk
from tkinter import filedialog

customtkinter.set_appearance_mode("dark")
customtkinter.set_default_color_theme("dark-blue")
root = customtkinter.CTk()
root.geometry("750x500")
root.minsize(width=650, height=300)
root.title('Whatnot Seller Tool - By Abandoned Treasures Reclaimed')
defEditor = "notepad.exe"



def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

root.iconbitmap(resource_path('16_3x_XIg_icon.ico'))

def get_all_names(np_arr):
    unique_names = []
    for i in range(0, len(np_arr)):
        unique_names.append(np_arr[i][2])
    unique_names = list(set(unique_names))
    unique_names.sort()
    return unique_names

def get_all_categories(np_arr):
    unique_categories = []
    for i in range(0, len(np_arr)):
        unique_categories.append(np_arr[i][22])
    unique_categories = list(set(unique_categories))
    unique_categories.sort()
    return unique_categories

def rename_columns(livestream_report):
    df = pandas.read_csv(livestream_report)
    df = df.rename(columns=({'buyer': 'Buyer'}))
    df = df.rename(columns=({'product name': 'product_name'}))
    df = df.rename(columns=({'cancelled or failed': 'isCanceled'}))
    df = df.rename(columns=({'sold price': 'sold_price'}))
    df[['Category', 'Number']] = df.product_name.str.split("#", expand=True)
    df.to_csv('temporary.csv', index=False)
    return df.to_numpy()

def get_shipment(name, category):
    try:                 # Can add Buyer, Category, sold_price, product_name, isCancelled
        query = ("SELECT Number "
                 + "FROM 'temporary.csv' WHERE isCanceled IS NULL AND Category ='" + category + "' AND Buyer='" + name + "' ORDER BY Category")
        shipment = duckdb.sql(query).df()

        if not shipment.empty:
            unsort_shipment = []
            cnt = 1
            for i in range(0, len(shipment)):
                if math.isnan(shipment.iat[i, 0]):
                    unsort_shipment.append(cnt)
                    cnt += 1
                else:
                    unsort_shipment.append(int(shipment.iat[i, 0]))
            unsort_shipment.sort()
            str_shipment = ""
            for i in range(0, len(unsort_shipment)):
                if i == 0:
                    str_shipment += str(unsort_shipment[i])
                else:
                    str_shipment += ("     " + str(unsort_shipment[i]))
            return str_shipment
        else:
            return -1
    except duckdb.duckdb.IOException:
        print("ERROR: File is open in other program")
    except duckdb.duckdb.ParserException:
        print("Error: Incorrect Parameter specified")

def get_all_shipments(names, categories):
    all_shipments = []
    for i in range(0, len(names)):
        shipments = []
        for j in range(0, len(categories)):
            str_shipment = get_shipment(names[i], categories[j])
            if str_shipment != -1:
                shipments.append([categories[j], str_shipment])
        all_shipments.append([names[i], shipments])
    return all_shipments

def create_doc(info):

    df = pandas.DataFrame(info)
    doc = docx.Document()

    style = doc.styles['Normal']
    font = style.font
    font.name = 'Arial'
    font.size = Pt(10)

    sections = doc.sections
    for section in sections:
        section.top_margin = Cm(1)
        section.bottom_margin = Cm(1)
        section.left_margin = Cm(3)
        section.right_margin = Cm(3)

    for i in range(0, len(df)):
        buyer = df.iat[i,0]
        shipment = df.iat[i, 1]
        heading = '___ ' + str(i+1) + '. ' + buyer
        doc.add_heading(heading, level=1)
        table = doc.add_table(rows=0, cols=2)
        table.TopPadding =Cm(20)
        cnt = 1
        for i in range(0, len(shipment)):
            category = shipment[i][0]
            numbers = shipment[i][1]
            cat_cells = table.add_row().cells
            cat_cells[0].text = 'Category: '
            cat_cells[1].text = str(category)
            num_cells = table.add_row().cells
            num_cells[0].text = 'Item #: '
            num_cells[1].text = str(numbers)

    doc.save('PRINT ME.docx')
    os.startfile(resource_path('PRINT ME.docx'))
    os.remove(resource_path('temporary.csv'))
    button.configure(text="Upload Livestream Report")

def upload_file():
    button.configure(text="Creating Your Document\n\nPlease Wait")
    file_path = filedialog.askopenfilename(filetypes=[(".csv Files", "*.csv")])
    button.configure(text="Upload Livestream Report")
    if file_path:
        file_name = file_path
        file = rename_columns(file_name)
        all_names = get_all_names(file)
        all_categories = get_all_categories(file)
        shipments = get_all_shipments(all_names, all_categories)
        create_doc(shipments)


contain = customtkinter.CTkScrollableFrame(master=root)
contain.pack(pady=0, padx=0, fill="both", expand=True)

web_logo = customtkinter.CTkImage(light_image=Image.open(resource_path('web_logo.png')),
                                  dark_image=Image.open(resource_path('web_logo.png')),
                                  size=(535, 27))

web_label = customtkinter.CTkLabel(contain, text="", image=web_logo)
web_label.pack(pady=30)

wn_logo = customtkinter.CTkImage(light_image=Image.open(resource_path('wn.png')),
                                  dark_image=Image.open(resource_path('wn.png')),
                                  size=(445, 45))

wn_label = customtkinter.CTkLabel(contain, text="", image=wn_logo)
wn_label.pack(pady=30)

atr_logo = customtkinter.CTkImage(light_image=Image.open(resource_path('toollogo.png')),
                                  dark_image=Image.open(resource_path('toollogo.png')),
                                  size=(575, 256))

atr_label = customtkinter.CTkLabel(contain, text="", image=atr_logo)
atr_label.pack(pady=30)

help_logo = customtkinter.CTkImage(light_image=Image.open(resource_path('help.png')),
                                   dark_image=Image.open(resource_path('help.png')),
                                   size=(568, 56))

free_logo = customtkinter.CTkImage(light_image=Image.open(resource_path('free.png')),
                                   dark_image=Image.open(resource_path('free.png')),
                                   size=(606, 107))

help_label = customtkinter.CTkLabel(contain, text="", image=help_logo)
help_label.pack(pady=10)

tool_label = customtkinter.CTkLabel(contain, text="", image=free_logo)
tool_label.pack(pady=20)

dono_logo = customtkinter.CTkImage(light_image=Image.open(resource_path('donate.png')),
                                   dark_image=Image.open(resource_path('donate.png')),
                                   size=(387, 63))

dono_label = customtkinter.CTkLabel(contain, text="", image=dono_logo)
dono_label.pack()

howto_logo = customtkinter.CTkImage(light_image=Image.open(resource_path('howto.png')),
                                    dark_image=Image.open(resource_path('howto.png')),
                                    size=(668, 324))

howto_label = customtkinter.CTkLabel(contain, text="", image=howto_logo)
howto_label.pack(pady=30)

button = customtkinter.CTkButton(master=contain, text="Upload Livestream Report", command=upload_file)
button.pack(pady=30, ipady=20, ipadx=10)

promo_logo = customtkinter.CTkImage(light_image=Image.open(resource_path('promo.png')),
                                    dark_image=Image.open(resource_path('promo.png')),
                                    size=(613, 127))

promo_label = customtkinter.CTkLabel(contain, text="", image=promo_logo)
promo_label.pack(pady=10)

made_logo = customtkinter.CTkImage(light_image=Image.open(resource_path('madeby.png')),
                                   dark_image=Image.open(resource_path('madeby.png')),
                                   size=(519, 56))

made_label = customtkinter.CTkLabel(contain, text="", image=made_logo)
made_label.pack(pady=10)

root.mainloop()