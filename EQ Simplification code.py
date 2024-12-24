import pandas as pd
from tkinter import filedialog
from customtkinter import*
from PIL import Image
from datetime import datetime
import os
import zipfile
import tkinter as tk
import warnings
from tkinter import messagebox
warnings.filterwarnings("ignore")
 
app = CTk()
 
app.title("GMI EQ")
set_appearance_mode("dark")
app.geometry("700x200")
 
EQ_CHILD_PATH = StringVar()
 
EQ_FILE_PATH = StringVar()
EQ_FILE_NAME = StringVar()
EQ_FILE = StringVar()
 
# XCAT FILES
EQ_FILE_XCAT_PATH = StringVar()
EQ_FILE_XCAT_NAME = StringVar()
EQ_FILE_XCAT = StringVar()
 
SAVE_FILE = StringVar()
 
#CHILD XCAT
def EQ_Files():
    EQ_LOCATION = filedialog.askdirectory()
    EQ_Entry.delete(0, tk.END)  
    EQ_Entry.insert(tk.END, EQ_LOCATION)
 
 
 
 
def EQ_Files_XCAT():
    EQ_XCAT_LOCATION = filedialog.askopenfilename(filetypes=[("ZIP files", "*.zip")])
    EQxcatfilename = os.path.split(EQ_XCAT_LOCATION)
    EQ_FILE_XCAT_PATH.set(EQxcatfilename[0]) #--- File path only
    EQ_FILE_XCAT_NAME.set(EQxcatfilename[1]) #--- File name only
    EQ_FILE_XCAT.set(EQ_XCAT_LOCATION) #--- Full file path and file name
 
def Output_file_eq():
    filenames = filedialog.askdirectory()  
    SAVE.delete(0, tk.END)  
    SAVE.insert(tk.END, filenames)
 
def EQ1():
   
    with zipfile.ZipFile(EQ_FILE_XCAT.get(), "r") as zip_ref:
        with zip_ref.open(zip_ref.namelist()[0]) as xcat:
            xcat = pd.read_excel(xcat)
 
    root_directory = EQ_CHILD_PATH.get()
    Child_path = []
 
    for root, dirs,files in os.walk(root_directory):
        for file in files:
               
            # Append the full path of the file to the list
            Child_path.append(os.path.join(root, file))
    print(Child_path)
    for child_ds in Child_path:
        with zipfile.ZipFile(child_ds, "r") as zip_ref:
               
                with zip_ref.open(zip_ref.namelist()[0]) as eq:
                    print("full",zip_ref.namelist()[0])
                    child = pd.read_excel(eq)
 
 
 
                    count_weight = child.columns.get_loc("WEIGHT")
                    merge = pd.merge(child, xcat, on ="ITM_ID", how = "left")
 
 
 
                    child.insert(count_weight+1, "WEIGHT XCAT",merge["WEIGHT_y"])
 
                    state = []
                    for ci in range(0,child.shape[0]):
                        if child["WEIGHT"][ci] == child["WEIGHT XCAT"][ci]:
                            state.append("TRUE")
                        else:
                            state.append("FALSE")
                           
                    child.insert(count_weight+2,"STATEMENT OF WEIGHT",state)
                    print(eq)
                    file_name = zip_ref.namelist()[0]
                    save_to_file_name =  SAVE_FILE.get() +"/" +   file_name.split("-")[1].upper()+".xlsx"
                    child = child.astype(str)
                    child.to_excel(save_to_file_name,index=False)
    messagebox.showinfo('Message',"Reports are formatted.")
 
 
 
EQ_LABEL = CTkLabel(master=app,text = "EQ CHILD",font=('Consolas',13.5),text_color="White")
EQ_LABEL.place(relx=0.03,rely=0.08)
 
EQ_Entry = CTkEntry(master=app,textvariable=EQ_CHILD_PATH,width=350,height=25,
                            text_color="White",placeholder_text="Extraction Report...",font=("Consolas",14))
EQ_Entry.place(relx=0.283,rely=0.08)
 
EQ_BUTTON = CTkButton(master=app,text=" Load  ",width=10,height=10,command=EQ_Files,corner_radius=45,fg_color="#FF6F00")
EQ_BUTTON.place(relx=0.81,rely=0.08)
 
 
 
EQ_LABEL_XCAT = CTkLabel(master=app,text = "EQ XCAT",font=('Consolas',13.5),text_color="White")
EQ_LABEL_XCAT.place(relx=0.03,rely=0.21)
 
EQ_Entry_XCAT = CTkEntry(master=app,textvariable=EQ_FILE_XCAT_NAME,width=350,height=25,
                            text_color="White",placeholder_text="Extraction Report...",font=("Consolas",14))
EQ_Entry_XCAT.place(relx=0.283,rely=0.21)
 
EQ_XCAT_BUTTON = CTkButton(master=app,text=" Load  ",width=10,height=10,command=EQ_Files_XCAT,corner_radius=45,fg_color="#FF6F00")
EQ_XCAT_BUTTON.place(relx=0.81,rely=0.21)
 
 
save_location = CTkLabel(master=app,text = "Filename to be Save",font=('Consolas',13.5),text_color="White")
save_location.place(relx=0.03,rely=0.32)
 
SAVE = CTkEntry(master=app,placeholder_text="Save file name...",textvariable=SAVE_FILE,width=350,height=25,
                            text_color="White",font=("Consolas",14))
SAVE.place(relx=0.283,rely=0.32)
 
 
EQ_SAVE_BUTTON = CTkButton(master=app,text=" Load  ",width=10,height=10,command=Output_file_eq,corner_radius=45,fg_color="#FF6F00")
EQ_SAVE_BUTTON.place(relx=0.81,rely=0.32)
 
 
generate = CTkButton(master=app,text="     Validate     ",width=20,height=10,command=EQ1,corner_radius=0)
generate.place(relx=0.81,rely=0.46)
 
 
 
 
 
 
 
 
app.mainloop()