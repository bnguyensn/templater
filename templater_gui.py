import tkinter as tk
from tkinter.filedialog import askopenfilename
from tkinter import ttk
from functools import partial
from templater import templater


class Application(ttk.Frame):
    def __init__(self, master=None):
        super().__init__(master)
        self.master = master
        self.pack()

        # Templater's necessary variables
        self.DATA_XLSX_PATH = ''
        self.TEMPLATE_DOCX_PATH = ''
        self.RESULT_FOLDER_PATH = ''

        # Set up GUI
        self.btnSelectDataFile = ttk.Button(text='Select data file',
                                            command=partial(self.select_file, [('Excel files', '*.xlsx')], 'DATA_XLSX_PATH'))
        self.btnSelectDataFile.pack()
        self.btnSelectWordFile = ttk.Button(text='Select template file',
                                            command=partial(self.select_file, [('Word files', '*.docx')], 'TEMPLATE_DOCX_PATH'))
        self.btnSelectWordFile.pack()
        self.btnRun = ttk.Button(text='Run', command=self.run)
        self.btnRun.pack()

        self.winfo_toplevel().title('Fuck Comcast v1.0')
        self.winfo_toplevel().update()
        self.winfo_toplevel().minsize(root.winfo_width() + 300, root.winfo_height() + 30)

    def select_file(self, ftypes, storagevar):
        fname = askopenfilename(title='Open', filetypes=ftypes)
        if fname:
            try:
                print('fname = {}'.format(fname))
                setattr(self, storagevar, fname)
            except:
                print('Error opening file.')
            return

    def run(self):
        print('DATA_XLSX_PATH = {}'.format(self.DATA_XLSX_PATH))
        print('TEMPLATE_DOCX_PATH = {}'.format(self.TEMPLATE_DOCX_PATH))
        templater.run(self.DATA_XLSX_PATH, self.TEMPLATE_DOCX_PATH)


root = tk.Tk()
app = Application(root)
app.mainloop()
