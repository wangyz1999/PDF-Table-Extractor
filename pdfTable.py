import tkinter as tk
from tkinter import filedialog
import pandas as pd
import pdfplumber as pp
from PIL import ImageTk, Image
import ctypes
import os

root = tk.Tk()
# root.geometry('300x300')
root.title("PDF Table")


class RectTracker:

    def __init__(self, canvas):
        self.canvas = canvas
        self.item = None

    def draw(self, start, end, **opts):
        """Draw the rectangle"""
        return self.canvas.create_rectangle(*(list(start) + list(end)), **opts)

    def autodraw(self, **opts):
        """Setup automatic drawing; supports command option"""
        self.start = None
        self.canvas.bind("<Button-1>", self.__start, '+')
        self.canvas.bind("<B1-Motion>", self.__update, '+')
        self.canvas.bind("<ButtonRelease-1>", self.__stop, '+')

        self._command = opts.pop('command', lambda *args: None)
        self.rectopts = opts

    def __start(self, event):
        global start_pos
        start_pos = [event.x, event.y]
        if not self.start:
            self.start = [event.x, event.y]
            return

        if self.item is not None:
            self.canvas.delete(self.item)
        self.item = self.draw(self.start, (event.x, event.y), **self.rectopts)
        self._command(self.start, (event.x, event.y))

    def __update(self, event):
        if not self.start:
            self.start = [event.x, event.y]
            return

        if self.item is not None:
            self.canvas.delete(self.item)
        self.item = self.draw(self.start, (event.x, event.y), **self.rectopts)
        self._command(self.start, (event.x, event.y))

    def __stop(self, event):
        global end_pos
        end_pos = [event.x, event.y]
        self.start = None
        self.canvas.delete(self.item)
        self.item = None
        self.draw(start_pos, end_pos, tags=('gre', 'box'))
        label_3["text"] = "Done Crop"


def get_file():
    file = filedialog.askopenfile()
    if not file:
        ctypes.windll.user32.MessageBoxW(0, "Must Choose PDF File", "Warning", 1)
        return
    global file_path
    file_path = file.name
    if file_path[-4:] != '.pdf':
        ctypes.windll.user32.MessageBoxW(0, "Must Choose PDF File", "Warning", 1)
        file_path = None
        return
    label_2["text"] = os.path.basename(file_path)
    label_3["text"] = ""
    label_4["text"] = ""


def crop_table():
    if not file_path:
        ctypes.windll.user32.MessageBoxW(0, "Choose a PDF file first", "Warning", 1)
        return
    pdf = pp.open(file_path)
    if entry1.get() == "":
        ctypes.windll.user32.MessageBoxW(0, "Choose a Page", "Warning", 1)
        return
    try:
        page_num = int(entry1.get()) - 1
    except:
        ctypes.windll.user32.MessageBoxW(0, "Enter an Integer", "Warning", 1)
        return
    try:
        page = pdf.pages[page_num]
    except:
        ctypes.windll.user32.MessageBoxW(0, "Page not Found", "Warning", 1)
        return

    label_3["text"] = ""
    label_4["text"] = ""

    IMG_PATH = "tmp_imgs/img.png"

    im = page.to_image()
    im.save(IMG_PATH, format="PNG")
    img = ImageTk.PhotoImage(Image.open(IMG_PATH))

    WIDTH = img.width()
    HEIGHT = img.height()

    im = page.to_image(resolution=288)

    im.save(IMG_PATH, format="PNG")

    image = Image.open(IMG_PATH)
    image = image.resize((WIDTH, HEIGHT), Image.ANTIALIAS)
    img = ImageTk.PhotoImage(image)

    WIDTH = img.width()
    HEIGHT = img.height()

    window = tk.Toplevel(root)
    canvas = tk.Canvas(window, width=WIDTH, height=HEIGHT)
    canvas.pack()
    canvas.create_image(0, 0, anchor='nw', image=img)
    canvas.image = img

    canv = canvas

    rect = RectTracker(canv)

    x, y = None, None

    def cool_design(event):
        global x, y
        kill_xy()

        dashes = [3, 2]
        x = canv.create_line(event.x, 0, event.x, 1000, dash=dashes, tags='no')
        y = canv.create_line(0, event.y, 1000, event.y, dash=dashes, tags='no')

    def kill_xy(event=None):
        canv.delete('no')

    canv.bind('<Motion>', cool_design, '+')

    rect.autodraw(fill="", width=1)


def preview_pdf():
    if not file_path:
        ctypes.windll.user32.MessageBoxW(0, "Choose a PDF file first", "Warning", 1)
        return
    pdf = pp.open(file_path)
    if entry1.get() == "":
        ctypes.windll.user32.MessageBoxW(0, "Choose a Page", "Warning", 1)
        return
    try:
        page_num = int(entry1.get()) - 1
    except:
        ctypes.windll.user32.MessageBoxW(0, "Enter an Integer", "Warning", 1)
        return
    try:
        page = pdf.pages[page_num]
    except:
        ctypes.windll.user32.MessageBoxW(0, "Page not Found", "Warning", 1)
        return
    if not start_pos or not end_pos:
        ctypes.windll.user32.MessageBoxW(0, "Crop Page First", "Warning", 1)
        return
    x0 = min(start_pos[0], end_pos[0])
    x1 = max(start_pos[0], end_pos[0])
    top = min(start_pos[1], end_pos[1])
    bottom = max(start_pos[1], end_pos[1])
    box = (x0, top, x1, bottom)
    page = page.within_bbox(box)

    if h.get() == 1:
        hori_stra = "lines"
    elif h.get() == 2:
        hori_stra = "text"
    if v.get() == 1:
        verti_stra = "lines"
    elif v.get() == 2:
        verti_stra = "text"

    dp = page.debug_tablefinder(table_settings={"vertical_strategy": verti_stra,
                                                "horizontal_strategy": hori_stra})
    page.to_image(resolution=150).outline_words()
    im = page.to_image(resolution=100)
    im.debug_tablefinder(tf=dp)
    PREVIEW_PATH = "tmp_imgs/preview.png"
    im.save(PREVIEW_PATH, format="PNG")
    img = ImageTk.PhotoImage(Image.open(PREVIEW_PATH))
    WIDTH = img.width()
    HEIGHT = img.height()

    window = tk.Toplevel(root)
    canvas = tk.Canvas(window, width=WIDTH, height=HEIGHT)
    canvas.pack()
    canvas.create_image(0, 0, anchor='nw', image=img)
    canvas.image = img


def save_to_excel():
    if not file_path:
        ctypes.windll.user32.MessageBoxW(0, "Choose a PDF file first", "Warning", 1)
        return
    pdf = pp.open(file_path)
    if entry1.get() == "":
        ctypes.windll.user32.MessageBoxW(0, "Choose a Page", "Warning", 1)
        return
    try:
        page_num = int(entry1.get()) - 1
    except:
        ctypes.windll.user32.MessageBoxW(0, "Enter an Integer", "Warning", 1)
        return
    try:
        page = pdf.pages[page_num]
    except:
        ctypes.windll.user32.MessageBoxW(0, "Page not Found", "Warning", 1)
        return
    if not start_pos or not end_pos:
        ctypes.windll.user32.MessageBoxW(0, "Crop Page First", "Warning", 1)
        return
    x0 = min(start_pos[0], end_pos[0])
    x1 = max(start_pos[0], end_pos[0])
    top = min(start_pos[1], end_pos[1])
    bottom = max(start_pos[1], end_pos[1])
    box = (x0, top, x1, bottom)
    page = page.within_bbox(box)
    if h.get() == 1:
        hori_stra = "lines"
    elif h.get() == 2:
        hori_stra = "text"
    if v.get() == 1:
        verti_stra = "lines"
    elif v.get() == 2:
        verti_stra = "text"
    table = page.extract_table(table_settings={"vertical_strategy": verti_stra,
                                               "horizontal_strategy": hori_stra})
    df = pd.DataFrame(table)
    e_filename = label_2["text"].split(".")[0] + "-p" + str(page_num + 1) + "."
    if r.get() == 1:
        e_filename += "xlsx"
    elif r.get() == 2:
        e_filename += "xls"
    df.to_excel(e_filename, index=None, header=None)
    label_4["text"] = "Done !"


start_pos, end_pos = None, None
file_path = None

button1 = tk.Button(root, text="Choose PDF File", command=get_file)
button1.grid(row=0, column=0, sticky='w')

label_2 = tk.Label(root, text="")
label_2.grid(row=0, column=1)

label_1 = tk.Label(root, text="Choose One Page:")
label_1.grid(row=1, column=0, sticky='w')

entry1 = tk.Entry(root)
entry1.grid(row=1, column=1)

button2 = tk.Button(root, text="Crop Table", command=crop_table)
button2.grid(row=2, column=0, sticky='w')

label_3 = tk.Label(root, text="")
label_3.grid(row=2, column=1)

label_5 = tk.Label(root, text="Horizontal:")
label_5.grid(row=3, column=0, sticky='w')

frame2 = tk.Frame(root)
frame2.grid(row=3, column=1)

h = tk.IntVar()
h.set(2)

radio3 = tk.Radiobutton(frame2, text="line", variable=h, value=1)
radio4 = tk.Radiobutton(frame2, text="text", variable=h, value=2)
radio3.pack(side="left")
radio4.pack(side="right")

label_6 = tk.Label(root, text="Vertical:")
label_6.grid(row=4, column=0, sticky='w')

frame2 = tk.Frame(root)
frame2.grid(row=4, column=1)

v = tk.IntVar()
v.set(2)

radio5 = tk.Radiobutton(frame2, text="line", variable=v, value=1)
radio6 = tk.Radiobutton(frame2, text="text", variable=v, value=2)
radio5.pack(side="left")
radio6.pack(side="right")

button3 = tk.Button(root, text="Preview Table", command=preview_pdf)
button3.grid(row=5, column=0, sticky='w')

button4 = tk.Button(root, text="Save to Excel", command=save_to_excel)
button4.grid(row=6, column=0, sticky='w')

r = tk.IntVar()
r.set(1)

frame1 = tk.Frame(root)
frame1.grid(row=6, column=1)

radio1 = tk.Radiobutton(frame1, text="xlsx", variable=r, value=1)
radio2 = tk.Radiobutton(frame1, text="xls", variable=r, value=2)
radio1.pack(side="left")
radio2.pack(side="right")

label_4 = tk.Label(root, text="")
label_4.grid(columnspan=7)


root.mainloop()
