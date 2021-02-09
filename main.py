from tkinter import filedialog
import tkinter.messagebox
from tkinter import ttk
from tkinter import *
from PyPDF2 import PdfFileReader, PdfFileWriter
from functools import partial
from win32com.shell import shell, shellcon

col = ["#000000", "#062f4f", "#813772", "#efefe8"]

def main():
    window = Tk()
    window.configure(bg=col[0])
    window.resizable(False,False)
    window.title("Merge/Split PDFs")
    s = ttk.Style()
    s.configure('TNotebook', font='Arial Bold', foreground='salmon', background=col[0])
    s.configure('TNotebook.Tab', background=col[1], foreground='black', lightcolor="#813772", borderwidth=0)
    tNotebook = ttk.Notebook(window)

    frame_merge = Frame(tNotebook, bg=col[1])
    Label(frame_merge,text="PDF list", bg=col[1], fg=col[3]).grid(row=0, columnspan=2,pady=10, padx=10)
    list_Of_Pdfs = Listbox(frame_merge, height="5", width="50",selectmode=EXTENDED)
    Button(frame_merge, text="Add", command=partial(browseFiles, list_Of_Pdfs)).grid(row=1, column=1,pady=10, padx=10, sticky="nsew")
    Button(frame_merge, text="Remove", command=partial(remove_From_List, list_Of_Pdfs)).grid(row=2, column=1,pady=10, padx=10, sticky="nsew")
    Button(frame_merge, text="Merge", command=partial(merge_pdfs, list_Of_Pdfs)).grid(row=3, columnspan=2,pady=10, padx=10,)
    list_Of_Pdfs.grid(row=1, rowspan=2,  pady=10, padx=10)

    frame_split = Frame(tNotebook, bg=col[1])
    Label(frame_split,text="", bg=col[1], fg=col[3]).grid(row=0,column=0, pady=10, padx=10)
    Label(frame_split,text="PDF path:", bg=col[1], fg=col[3]).grid(row=1,column=0, pady=10, padx=10)
    entry_Path = Entry(frame_split, width=40); entry_Path.grid(row=1, column=1, pady=10, padx=0, sticky="nsew")
    Button(frame_split,text="Browse", command=partial(browseFile, entry_Path)).grid(row=1, column=2, pady=10, padx=10)
    Label(frame_split,text="Pages num:", bg=col[1], fg=col[3]).grid(row=2,column=0, pady=10, padx=10)
    entry_Pages = Entry(frame_split, width=40); entry_Pages.grid(row=2, column=1, pady=10, padx=0, sticky="nsew")
    Button(frame_split, text="Split", command=partial(split_pdf, entry_Path, entry_Pages)).grid(row=3, columnspan=3, pady=10, padx=10,)

    tNotebook.add(frame_merge, text="Merge", padding=4, sticky="nsew")
    tNotebook.add(frame_split, text="Split", padding=4, sticky="nsew")
    tNotebook.pack(fill=BOTH, expand=True)
    window.mainloop()


def browseFiles(list_Of_Pdfs):
    filename = filedialog.askopenfilename(initialdir = "/", title = "Select a File", multiple=True,
                                            filetypes = (("Pdf files","*.pdf*"),("all files","*.*")))
    if filename != []:
        for file in filename:
            list_Of_Pdfs.insert(list_Of_Pdfs.size(),file)


def browseFile(entry_Path):
    filename = filedialog.askopenfilename(initialdir = "/", title = "Select a File",
                                            filetypes = (("Pdf files","*.pdf*"),("all files","*.*")))
    if filename != "":
        entry_Path.insert(0, filename)


def remove_From_List(list_Of_Pdfs):
    for idx in list_Of_Pdfs.curselection():
        list_Of_Pdfs.delete(idx)


def onselect(evt):
    return evt.widget.curselection()


def merge_pdfs(list_Of_Pdfs, output= shell.SHGetFolderPath (0, shellcon.CSIDL_DESKTOP, 0, 0) + r"\Merged.pdf" ):
    pdf_writer = PdfFileWriter()

    if list_Of_Pdfs.size() != 0:
        for path in list_Of_Pdfs.get(0,list_Of_Pdfs.size()):
            pdf_reader = PdfFileReader(path)
            for page in range(pdf_reader.getNumPages()):
                pdf_writer.addPage(pdf_reader.getPage(page))
        with open(output, 'wb') as out:
            pdf_writer.write(out)
    messagebox.showinfo(title="Created", message="Pdf is created successfully. :)")

def split_pdf(entry_Path, entry_Pages):
    path = entry_Path.get()
    pages_split = entry_Pages.get().strip()
    if path and pages_split != "":
        pdf = PdfFileReader(path)
        pages_num = []
        for seq in pages_split.split(","):
            if seq.isnumeric():
                pages_num.append(int(seq))
            elif seq != "":
                num1,num2 = seq.strip().split("-")
                for num in range(int(num1)-1, int(num2)):
                    pages_num.append(int(num))
    pdf_writer = PdfFileWriter()
    for page in pages_num:
        pdf_writer.addPage(pdf.getPage(page))
    output = shell.SHGetFolderPath (0, shellcon.CSIDL_DESKTOP, 0, 0) + r"\Splitted.pdf"
    with open(output, 'wb') as output_pdf:
        pdf_writer.write(output_pdf)
    messagebox.showinfo(title="Created", message="Pdf is created successfully. :)")

main()
