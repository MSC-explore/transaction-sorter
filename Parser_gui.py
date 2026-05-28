import customtkinter
from customtkinter import *
from tkinter import ttk
from pathlib import Path
#importing parse_pdf and save_File functions
from Bank_info_Parsing2 import *

customtkinter.set_appearance_mode("dark")
customtkinter.set_default_color_theme("green")
docs_folder = Path.home() / "Documents"

def openFile():
    filename =  filedialog.askopenfilename(title="Open PDF file",
                                           initialdir=docs_folder,
                                           filetypes=[('PDF files', '*.pdf')])

    filename_label.configure(text=filename)

def save_location():
    savefilename =  filedialog.asksaveasfilename(title="Choose Save Location",
                                           initialdir=docs_folder,
                                           filetypes=[('Excel files', '*.xlsx')])

    savefilename_label.configure(text=savefilename)

def clear_File():
    filename_label.configure(text="No File Selected")
    savefilename_label.configure(text="No File Selected")
    preview_xl.delete(*preview_xl.get_children())
    savepreview_label.configure(text="    ", text_color="green")

def parseNload():
    duplicates.update(parse_pdf(filename_label.cget("text")))
    preview_xl.delete(*preview_xl.get_children())
    for desc, info in duplicates.items():
        preview_xl.insert("", "end", values=(
            desc,
            info["count"],
            round(info["total"], 2),
            round(info["total"] / info["count"], 2),
            ", ".join(f"{a:.2f}" for a in info["amounts"]),
        ))

def Save_Message():
    savepreview_label.configure(text="Saved!", text_color="green")
    
def saveNmessage():
    save_File(duplicates, savefilename_label.cget("text"))
    Save_Message()

headers = ["Description", "Times Repeated", "Total Amount", "Average Amount", "All Amounts"]

duplicates={}

#parent widget (the container)
root = customtkinter.CTk()
root.geometry("1000x800")
root.title("PDF Text Extractor")

chooseframe = CTkFrame(root, fg_color="transparent")
saveframe = CTkFrame(root, fg_color="transparent")
parseframe = CTkFrame(root, fg_color="transparent")
savepreviewframe = CTkFrame(root, fg_color="transparent")
clearframe =CTkFrame(root, fg_color="transparent")

#border and title
title_frame = CTkFrame(root, border_width=2, border_color="white", fg_color="transparent")
title_label = CTkLabel(title_frame, text="Transaction Sorter", font=("FixedSys", 40))

description_label = CTkLabel(title_frame, text="(Find Recurring Purchases from your USAA pdf)")
choosefile_button = CTkButton(chooseframe, text="Choose a PDF File", command=openFile)
filename_label = CTkLabel(chooseframe, text="No File Selected")
savefilename_label = CTkLabel(saveframe, text="No Location Selected")
saveas_button = CTkButton(saveframe, text="Choose a Save Location", command=save_location)
parsefile_button = CTkButton(parseframe, text="Find Duplicate Transactions", command=parseNload)
Preview_label = CTkLabel(parseframe, text="Preview", font=CTkFont(size=20))
savepreview_button = CTkButton(savepreviewframe, text="Save Preview", command=saveNmessage) 
savepreview_label = CTkLabel(savepreviewframe, text="      ", text_color="green")
clearfile_button = CTkButton(clearframe, text="Clear Files", command=clear_File)
preview_xl = ttk.Treeview(parseframe, columns=headers, show="headings")

# changing the preview color
style = ttk.Style()
style.theme_use("clam")
style.layout("Treeview", [
    ("Treeview.treearea", {"sticky": "nswe"})
])
# body of the spreadsheet
style.configure("Treeview", 
                foreground="white", 
                background="black", 
                fieldbackground="Black",
                borderwidth=0,
                relief="flat")
# headers colors
style.configure("Treeview.Heading", 
                background="black", 
                foreground="white")

# Highlight Colors
style.map("Treeview", background=[("selected", "#333333")])
style.map("Treeview.Heading", background=[("selected", "#333333")])


for header in headers:
    preview_xl.heading(header, text= header)


title_label.pack(pady=10)
description_label.pack()
choosefile_button.pack(side=LEFT, padx=20)
filename_label.pack(side=RIGHT)


saveas_button.pack(side=LEFT, padx=10)
savefilename_label.pack(side=RIGHT)

parsefile_button.pack()
Preview_label.pack(pady=10)
preview_xl.pack(expand=True, fill=BOTH)
savepreview_button.pack(side=LEFT, padx=10)
savepreview_label.pack(side=RIGHT)

clearfile_button.pack()

title_frame.pack()
chooseframe.pack(pady=10)
parseframe.pack(pady=10, expand=True, fill=BOTH)
saveframe.pack(pady=10)
savepreviewframe.pack(pady=10)
clearframe.pack(pady=10)

root.mainloop()