from cryptography.fernet import Fernet
import base64
import customtkinter
from PIL import Image, ImageTk
from tkinter import filedialog
import csv
import os
import sys
import subprocess as sp

code = b"""


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
def atr_tool(path):
    username = []
    prod_name = []
    print_list=""
    with open(path, mode='r')as file:
        csvFile = csv.reader(file)
        next(csvFile)
        for lines in csvFile:
            if len(lines[7]) == 0:
                username.append(lines[2])
                prod_name.append(lines[3])

    uni_username = list(set(username))

    in_arr = []
    out_arr = []
    p = 0
    for i in range(0, len(uni_username)):
        for j in range(0, len(username)):
            if uni_username[i] == username[j]:
                if p == 0:
                    in_arr.append(uni_username[i])
                    p = 1
                s = prod_name[j]
                s1 = s.split('#')           # Split at # because it seperates the weight(category) from auction number
                if len(s1) < 2:             # If it didnt find a # because it was a givey or single item then assign it 8888 for seller notice
                    s1.append('8888')
                in_arr.append(s1)
        out_arr.append(in_arr)
        in_arr = []
        p = 0

    prod_category = []                      # Gathers all product weights/categories from a sellers show
    for i in range(0, len(out_arr)):
        for j in range(0, len(out_arr[i])):
            if len(out_arr[i][j]) == 2:
                prod_category.append(out_arr[i][j][0])

    uni_prod_category = list(set(prod_category))

    final_username = []
    final_prod_category = []
    for i in range(0, len(uni_username)):
        final_username.append(uni_username[i])
        final_username.append(uni_prod_category)

    final_list = []
    current_category = []
    current_items = []

    for i in range(0, len(uni_username)):           # Creates a master 3D Array with all unique buyers, all unique weights/categories,
        final_list.append(uni_username[i])          # and each individual sale per buyer but also but also leaves blanks for categories with no purchase
        for j in range(0, len(uni_prod_category)):
            current_category.append(uni_prod_category[j])
            current_username = ""
            for x in range(0, len(out_arr)):
                for y in range(0, len(out_arr[x])):
                    if len(out_arr[x][y]) != 2:
                        current_username = out_arr[x][y]
                    else:
                        if current_username == uni_username[i]:
                            if uni_prod_category[j] == out_arr[x][y][0]:        # Converts string numbers to int so that they may be sorted
                                current_items.append(int(out_arr[x][y][1]))
            current_items.sort()
            for x in range(0, len(current_items)):                              # Then converts them back to string for output
                current_items[x] = str(current_items[x])
            current_category.append(current_items)
            current_items = []
        final_list.append(current_category)
        current_category = []

    temp_category = ""
    tmp = 0
    print_list = []
    print_category = []

    for i in range(0, len(final_list)):         # Generates the simple to read printable list that removes the blanks from the master list
        if i % 2 == 0:                          # that would occur when a buyer did not purchase a product in a certain category
            print_list.append(final_list[i])
        else:
            for j in range(0, len(final_list[i])):
                if j % 2 != 0:
                    if len(final_list[i][j]) > 0:
                        if tmp == 0:
                            print_category.append(temp_category)
                            tmp = 1
                        print_category.append(final_list[i][j])
                else:
                    temp_category = final_list[i][j]
                tmp = 0
        if len(print_category) > 0:
            print_list.append(print_category)
            print_category = []
    return print_list


def final_string(input):
    final_str = ""
    for i in range(0, len(input)):
        if i % 2 == 0:
            final_str += '--------------------------------------------------------------------------------'
            final_str += '\\n\'
            final_str += ('___ ' + input[i] + '\\n\' + '\\n\')

        else:
            for j in range(0, len(input[i])):
                if j % 2 == 0:
                    final_str += ('Category:\t\t' + input[i][j] + '\\n\')
                else:
                    final_str += 'Item #:   \t\t'
                    for k in range(0, len(input[i][j])):
                        final_str += (input[i][j][k] + '   ')
                    final_str += '\\n\'
                    if j != (len(input[i])-1):
                        final_str += '\\n\'
            final_str += '\\n\'
            final_str += '\\n\'
    final_str += '--------------------------------------------------------------------------------'
    final_str += '\\n\'
    final_str += '\\n\'
    return final_str

def create_doc(info):
	file_path = 'PRINT ME.txt'
	with open(file_path, 'w') as file:
		file.write(info)
	sp.Popen([defEditor, file_path])

def upload_file():
    file_path = filedialog.askopenfilename(filetypes=[(".csv Files", "*.csv")])
    if file_path:
        tmp = atr_tool(file_path)
        print(tmp)
        final_list = final_string(tmp)
        create_doc(final_list)


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
"""

key = Fernet.generate_key()
encryption_type = Fernet(key)
encrypted_message = encryption_type.encrypt(code)
decrypted_message = encryption_type.decrypt(encrypted_message)

decrypted_message.replace(b'|', b'\n')
exec(decrypted_message)
