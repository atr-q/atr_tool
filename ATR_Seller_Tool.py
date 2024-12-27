import duckdb, pandas
import math, numpy
import os, sys, errno
import subprocess as sp
import docx
from docx.shared import Pt
from docx.shared import Cm
import customtkinter
from PIL import Image
from tkinter import filedialog

customtkinter.set_appearance_mode("dark")
customtkinter.set_default_color_theme("green")
root = customtkinter.CTk()
root.geometry("1280x750")
root.minsize(width=600, height=400)
root.title('Whatnot Seller Tool - By Abandoned Treasures Reclaimed')
pf = os.getenv("ProgramFiles")
wordEditor = pf + "\\Windows NT\\Accessories\\wordpad.exe"
defEditor = "notepad.exe"
radio_var = customtkinter.StringVar(value="n")


def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

root.iconbitmap(resource_path('16_3x_XIg_icon.ico'))

def silent_remove(filename):
    try:
        os.remove(resource_path(filename))
    except OSError as e: # this would be "except OSError, e:" before Python 2.6
        if e.errno != errno.ENOENT: # errno.ENOENT = no such file or directory
            raise # re-raise exception if a different error occurred

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

def create_word_doc(info):

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
        heading = str(i+1) + '.___ ' + buyer
        doc.add_heading(heading, level=1)
        table = doc.add_table(rows=0, cols=2)
        for j in range(0, len(shipment)):
            category = shipment[j][0]
            numbers = shipment[j][1]
            cat_cells = table.add_row().cells
            cat_cells[0].text = 'Category: '
            cat_cells[1].text = str(category)
            num_cells = table.add_row().cells
            num_cells[0].text = 'Item #: '
            num_cells[1].text = str(numbers)
    filename = 'PRINT ME.docx'
    if radio_var.get() == 'w':
        silent_remove(filename)
        doc.save(filename)
        sp.Popen([wordEditor, filename])
    elif radio_var.get() == 'm':
        doc.save(filename)
        os.startfile(resource_path(filename))
    os.remove(resource_path('temporary.csv'))
    button.configure(text="Upload Livestream Report")

def create_txt_doc(info):
    df = pandas.DataFrame(info)
    final_str = ""
    for i in range(0, len(df)):
        buyer = df.iat[i,0]
        shipment = df.iat[i, 1]
        final_str += '--------------------------------------------------------------------------------\n'
        final_str += str(i+1) + '.___ ' + buyer + '\n'
        for j in range(0, len(shipment)):
            category = shipment[j][0]
            numbers = shipment[j][1]
            final_str += ('Category:\t\t' + category + '\n')
            final_str += ('Item #:   \t\t' + numbers + '\n')
            final_str += '\n\n'
    file_path = 'PRINT ME.txt'
    with open(file_path, 'w') as file:
        file.write(final_str)
    sp.Popen([defEditor, file_path])

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
        if radio_var.get() == 'm' or radio_var.get() == 'w':
            create_word_doc(shipments)
        elif radio_var.get() == 'n':
            create_txt_doc(shipments)



contain = customtkinter.CTkScrollableFrame(master=root, fg_color='#0E0C07')
contain.pack(fill="both", expand=True)


gui = customtkinter.CTkImage(light_image=Image.open(resource_path('gui.png')),
                                  dark_image=Image.open(resource_path('gui.png')),
                                  size=(1280, 520))

gui_label = customtkinter.CTkLabel(contain, text="", image=gui)
gui_label.pack()

doc_choice_notepad = customtkinter.CTkRadioButton(contain, text="Notepad Document", value='n', fg_color='#D6B44A', hover_color='white', variable=radio_var)
doc_choice_notepad.pack(pady=5)

doc_choice_wordpad = customtkinter.CTkRadioButton(contain, text="Wordpad Document", value='w', fg_color='#D6B44A', hover_color='white', variable=radio_var)
doc_choice_wordpad.pack(pady=5)

doc_choice_docx = customtkinter.CTkRadioButton(contain, text="Microsoft Word Document", value='m', fg_color='#D6B44A', hover_color='white', variable=radio_var)
doc_choice_docx.pack(pady=5)

button = customtkinter.CTkButton(master=contain, text="Upload Livestream Report", fg_color='#D6B44A', hover_color='white', text_color='black', command=upload_file)
button.pack(pady=30, ipady=20, ipadx=10)

root.mainloop()