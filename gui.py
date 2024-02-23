from tkinter import *
from tkinter import filedialog
from tkinter import ttk

def main():
    global root, files
    root = Tk()
    root.title('Speadsheet Slicer-dicer')

    files = {'subject': StringVar(), 'email': StringVar(), 'css': StringVar(), 'csv': StringVar(), 'attachment': StringVar(), 'test': BooleanVar(value=True)}
    set_defaults()

    mainframe = ttk.Frame(root, padding='3 3 12 12')
    root.columnconfigure(0, weight=5)
    root.rowconfigure(0, weight=5)

    # Text field
    txt_subject     = ttk.Entry(mainframe, textvariable=files['subject'])

    # Buttons
    btn_email       = ttk.Button(mainframe, text='Browse', command=lambda: browse_button('email'))
    btn_css         = ttk.Button(mainframe, text='Browse', command=lambda: browse_button('css'))
    btn_csv         = ttk.Button(mainframe, text='Browse', command=lambda: browse_button('csv'))
    btn_attachment  = ttk.Button(mainframe, text='Browse', command=lambda: browse_button('attachment'))
    btn_send        = ttk.Button(mainframe, text='Send Emails', command=send)

    # Checkbutton
    cbt_testing     = ttk.Checkbutton(mainframe, text='Test: Don\'t actually send the email', variable=files['test'], onvalue=True, offvalue=False)

    # Bindings
    btn_send    .focus_set()

    # Labels
    lbl_title           = ttk.Label(mainframe, text='Files')
    lbl_email           = ttk.Label(mainframe, text='Email Template')
    lbl_css             = ttk.Label(mainframe, text='CSS File')
    lbl_csv             = ttk.Label(mainframe, text='CSV File')
    lbl_attachment      = ttk.Label(mainframe, text='Attachment')
    lbl_file_email      = ttk.Label(mainframe, textvariable=files['email'])
    lbl_file_css        = ttk.Label(mainframe, textvariable=files['css'])
    lbl_file_csv        = ttk.Label(mainframe, textvariable=files['csv'])
    lbl_file_attachment = ttk.Label(mainframe, textvariable=files['attachment'])

    # Positions
    mainframe           .grid(column=0, row=0, sticky=(N, W, E, S))

    lbl_title           .grid(column=1, row=1, sticky=W)

    txt_subject         .grid(column=1, row=2, sticky=N+S+E+W, columnspan=8)

    lbl_email           .grid(column=1, row=3, sticky=W)
    btn_email           .grid(column=2, row=3, sticky=W)
    lbl_file_email      .grid(column=3, row=3, sticky=W)

    lbl_css             .grid(column=1, row=4, sticky=W)
    btn_css             .grid(column=2, row=4, sticky=W)
    lbl_file_css        .grid(column=3, row=4, sticky=W)

    lbl_csv             .grid(column=1, row=5, sticky=W)
    btn_csv             .grid(column=2, row=5, sticky=W)
    lbl_file_csv        .grid(column=3, row=5, sticky=W)

    lbl_attachment      .grid(column=1, row=6, sticky=W)
    btn_attachment      .grid(column=2, row=6, sticky=W)
    lbl_file_attachment .grid(column=3, row=6, sticky=W)

    cbt_testing         .grid(column=1, row=7, sticky=W)

    btn_send            .grid(column=1, row=8, sticky=W)

    for child in mainframe.winfo_children():
        child.grid_configure(padx=5, pady=5)

    root.mainloop()

def browse_button(file):
    """folder path = StringVar()"""
    files[file].set(filedialog.askopenfile(
        title=f'Select {file}',
        filetypes=(('All files', '*.*'),)
        ).name)

def set_defaults():
    import emailer
    files['subject'].set(emailer.subject)
    files['email'].set(emailer.email_file)
    files['css'].set(emailer.css_file)
    files['csv'].set(emailer.csv_file)
    files['attachment'].set(emailer.attachment_file)

def send(event=None):
    for key, var in files.items():
        files[key] = var.get()
    root.destroy()

def flip(key):
    files[key].set(not files[key].get())

def inspect_gui(func):
    def decorator(*args, **kwargs):
        print(*((k, v.get()) for k, v in files.items()), sep='\n')
        func(*args, **kwargs)
    return decorator

if __name__ == '__main__':
    send = inspect_gui(send)  # Decorate send(): Inspect inputs collected by GUI.
    main()
