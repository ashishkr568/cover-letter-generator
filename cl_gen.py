"""
Code to update fields such as Company, Location, Date and Role in a standard cover letter template and
save it in pdf format.

App Name: Cover Letter Generator
Version: 1.0
Author: Ashish Kumar
Date: 08-Feb-2022
LinkedIn: : https://www.linkedin.com/in/ashish568/
Medium: https://medium.com/@ashish.568
GitHub: https://github.com/ashishkr568

-- Creation of .exe (using the spec file) --
pyinstaller --onefile cl_gen.spec

"""

import tkinter as tk
import tkinter.filedialog
import webbrowser
from datetime import date

# import packages
from fpdf import FPDF

# Create a Tkinter App, set canvas size and set its icon and title
app = tk.Tk()
app.geometry("500x450")
# app.configure(background='#0077b5')
app.iconbitmap("img/appIcon.ico")
app.title("Cover Letter Generator")
# app.eval('tk::PlaceWindow . center')
app.resizable(False, False)


# All required Functions


# Function to Check if all values are filled
def on_generate_cl_click():
    val_dict = fetch_values()
    if val_dict["inp_path"] == "" or val_dict["out_dir"] == "" or val_dict["new_f_name"] == "" \
            or val_dict["company"] == "" or val_dict["location"] == "" or val_dict["position"] == "":
        if val_dict["inp_path"] == "":
            tk.messagebox.showerror('Error', 'Error: Please select cover letter template..!!')
        # Check if Output Path is empty
        if val_dict["out_dir"] == "":
            tk.messagebox.showerror('Error', 'Error: Please select output path..!!')

        # Check if Output File Name is empty
        if val_dict["new_f_name"] == "":
            tk.messagebox.showerror('Error', 'Error: Please enter output file name..!!')
            # Check if Output File Name is empty
        if val_dict["company"] == "":
            tk.messagebox.showerror('Error', 'Error: Please enter company name..!!')
            # Check if Output File Name is empty
        if val_dict["location"] == "":
            tk.messagebox.showerror('Error', 'Error: Please enter job location..!!')
            # Check if Output File Name is empty
        if val_dict["position"] == "":
            tk.messagebox.showerror('Error', 'Error: Please enter role/position you are applying for..!!')
    else:
        val_dict = fetch_values()
        generate_cl(val_dict)
        tk.messagebox.showinfo("Message", "Cover Letter Generated..!!")

        # Delete all filled values from form
        e1.delete("0", "end")
        e2.delete("0", "end")
        e3.delete("0", "end")
        e5.delete("0", "end")
        e6.delete("0", "end")
        e7.delete("0", "end")


# Function to fetch values from Form filled on Tkinter app
def fetch_values():
    inp_path = str(e1.get())
    out_dir = str(e2.get())

    # Check if User has provided .pdf in file name, If yes then strip it
    new_f_name = e3.get()
    new_f_name = new_f_name.replace(".pdf", "")

    app_company = str(e5.get())
    app_location = str(e6.get())
    app_position = str(e7.get())

    # Add these values to a Dictionary
    val_dict = {"inp_path": inp_path,
                "out_dir": out_dir,
                "new_f_name": new_f_name,
                "company": app_company,
                "location": app_location,
                "position": app_position}
    return val_dict


# Function to update cover letter and save output to a pdf file
def generate_cl(val_dict):
    # Read template file from the Directory
    cl_template = open(val_dict["inp_path"], 'r')
    template_text = cl_template.read()

    # Update cover letter field values
    company = val_dict["company"]
    location = val_dict["location"]
    app_date = str(date.today().strftime("%d %b %Y"))
    role = val_dict["position"]

    # Replace fields in template file
    variables = {"company": company, "location": location, "date": app_date, "role": role}
    updated_cl = template_text.format(**variables)

    # # Save updated file as .txt
    # with open('readme.txt', 'w') as f:
    #     f.write(updated_cl)
    #     f.write('\n')

    # Save updated Cover letter in PDF format

    # This example makes use of the header and footer methods to process page headers and footers.
    # They are called automatically. They already exist in the FPDF class but do nothing,
    # therefore we have to extend the class and override them.
    class PDF(FPDF):
        def header(self):
            # Arial bold 15
            self.set_font('Arial', 'B', 20)
            # Move to the right
            self.cell(75)
            # Title
            self.cell(30, 10, 'Cover Letter', 'C')
            # Line break
            self.ln(20)

    # Extend the PDF Class
    pdf = PDF()

    # Add a page
    pdf.add_page()

    # set style and size of font for pdf body
    pdf.set_font("Arial", size=10)

    # open the text file in read mode (In case we need to create pdf from content of a text file
    # f = open("cl_template.txt", "r")
    # # insert the texts in pdf
    # for x in f:
    #     pdf.multi_cell(193, 5, txt=x, align='l')

    # Write updated content to pdf
    pdf.multi_cell(193, 5, txt=updated_cl, align='l')

    # Create a box to enclose content
    pdf.rect(x=5, y=5, w=200, h=287, style='')

    # Save File at output location
    try:
        pdf.output(val_dict["out_dir"] + '/' + val_dict["new_f_name"] + '.pdf')
    except PermissionError:
        tk.messagebox.showerror('Error',
                                'Error: There is already an existing file with the same name and its open,'
                                ' Please Close it and try again..!!')
        val_dict = fetch_values()
        generate_cl(val_dict)
        tk.messagebox.showinfo("Message", "Cover Letter Generated..!!")

        # Delete all filled values from form
        e1.delete("0", "end")
        e2.delete("0", "end")
        e3.delete("0", "end")
        e5.delete("0", "end")
        e6.delete("0", "end")
        e7.delete("0", "end")


# Create Form in Tkinter App using Grid

# Row = 0

# Create a Frame to enter App Description -- Column = All
app_intro = "This Tool will update Company Name, Job Location and Job Title in a Standard Cover Letter Template File"

# Create App Description Label
app_intro_frame = tk.Frame(app, width=400, height=70)
app_intro_frame.grid(row=0, columnspan=3, padx=10, pady=5, ipadx=1, ipady=3, sticky='news')
# app_intro_frame.grid_propagate(False)
# app_intro_frame.update()
app_intro_label = tk.Label(app_intro_frame, text=app_intro, wraplength=400, font=("Arial Bold", 8, "italic"))
app_intro_label.place(x=250, y=30, anchor="center")

# Take File Information from User

# Row = 1

# Input File Location

# Define Label -- Column = 0
tk.Label(app, text="Cover Letter Template").grid(row=1, column=0, ipadx=10, ipady=5, sticky='news')

# Define TextBox -- Column = 1
e1 = tk.Entry(app, state='disabled')
e1.grid(row=1, column=1, padx=10, pady=5, ipadx=1, ipady=3, sticky='news')


# Define a Browse Button for Input File Selection -- column = 2
# Function to open a new dialog and select pdf file
def inp_browse():
    inp_filename = tkinter.filedialog.askopenfiles(title="Select a File", filetypes=(("text Files", "*.txt"),))
    # print(str(inp_filename[0].name))
    e1.config(state='normal')
    e1.insert(0, str(inp_filename[0].name))
    e1.config(state='disabled')


# Function to Update color on hover for input browse button
def inp_button_on_enter(e):
    inp_browse_button['background'] = 'Black'
    inp_browse_button['foreground'] = 'white'


def inp_button_on_leave(e):
    inp_browse_button['background'] = 'SystemButtonFace'
    inp_browse_button['foreground'] = 'black'


# Create Input Browse Button
inp_browse_button = tkinter.Button(app, text="Browse", command=inp_browse)
inp_browse_button.grid(row=1, column=2, sticky='news', padx=10, pady=5)

# Add Mouse Hover Effect
inp_browse_button.bind("<Enter>", inp_button_on_enter)
inp_browse_button.bind("<Leave>", inp_button_on_leave)

# Row = 2

# Output File Path -- Column = 0
# Define Label -- Column = 0
tk.Label(app, text="Output Path").grid(row=2, column=0, ipadx=10, ipady=5, sticky='news')
# Define TextBox -- Column = 1
e2 = tk.Entry(app, state='disabled')
e2.grid(row=2, column=1, padx=10, pady=5, ipadx=1, ipady=3, sticky='news')


# Define a Browse Button To fetch Output Directory Location
# Function to Open a dialog for selecting output location
def out_dir_browse():
    out_dir = tkinter.filedialog.askdirectory(title="Select Output Directory")
    # print(out_dir)
    e2.config(state='normal')
    e2.insert(0, str(out_dir))
    e2.config(state='disabled')


# Function to Update color on hover for Output browse button
def out_button_on_enter(e):
    out_browse_button['background'] = 'Black'
    out_browse_button['foreground'] = 'white'


def out_button_on_leave(e):
    out_browse_button['background'] = 'SystemButtonFace'
    out_browse_button['foreground'] = 'black'


# Create Button for selecting output location
out_browse_button = tkinter.Button(app, text="Browse", command=out_dir_browse)
out_browse_button.grid(row=2, column=2, sticky='news', padx=10, pady=5)

# Add Mouse Hover Effect
out_browse_button.bind("<Enter>", out_button_on_enter)
out_browse_button.bind("<Leave>", out_button_on_leave)

# Row = 3

# Output File Name
# Define Label -- Column = 0
tk.Label(app, text="Output File Name").grid(row=3, column=0, ipadx=10, ipady=5, padx=10, pady=5, sticky='news')
# Define TextBox -- Column = 1
e3 = tk.Entry(app)
e3.grid(row=3, column=1, padx=10, pady=5, ipadx=1, ipady=3, sticky='news')

# Row = 4

# Create a frame to include text -- Column = All
app_frame = tk.Frame(app, width=400, height=50)
app_frame.grid(row=4, column=0, columnspan=3)
app_frame.grid_propagate(False)
app_frame.update()
enter_details_label = tk.Label(app_frame, text="Enter Details", wraplength=400,
                               font=("Arial Bold", 8, "italic", "underline"))
enter_details_label.place(x=200, y=30, anchor="center")

# Row = 5

# Company Name
# Define Label -- Column = 0
tk.Label(app, text="Company Name").grid(row=5, column=0, ipadx=10, ipady=5, sticky='news')
# Define TextBox -- Column = 1
e5 = tk.Entry(app)
e5.grid(row=5, column=1, padx=10, pady=5, ipadx=1, ipady=3, sticky='news')

# Row = 6

# Job Location
# Define Label -- Column = 0
tk.Label(app, text="Job Location").grid(row=6, column=0, ipadx=10, ipady=5, sticky='news')
# Define TextBox -- Column = 1
e6 = tk.Entry(app)
e6.grid(row=6, column=1, padx=10, pady=5, ipadx=1, ipady=3, sticky='news')

# Row = 7

# Position
# Define Label -- Column = 0
tk.Label(app, text="Job Role/Position").grid(row=7, column=0, ipadx=10, ipady=5, sticky='news')
# Define TextBox -- Column = 1
e7 = tk.Entry(app)
e7.grid(row=7, column=1, padx=10, pady=5, ipadx=1, ipady=3, sticky='news')

# Row = 8

# Close and Generate Cover Letter Buttons in the same Grid Cell using Frame
# Define Frame
f1 = tk.Frame(app)
f1.grid(row=8, column=1, sticky="nsew", columnspan=2, padx=10, pady=5, ipadx=1, ipady=3)
# Add close button with image -- column = 1
close_button_img = tk.PhotoImage(file="img/close_button.png")
c_button = tk.Button(f1, text="Exit Application", image=close_button_img, borderwidth=0, cursor='hand2',
                     command=app.quit)

# Add "Generate Cover Letter" Button with color change on hover
run_b2 = tk.Button(f1, text='Generate Cover Letter', cursor='hand2', command=on_generate_cl_click)


# Color change on hover for Generate Cover Letter Button
def generate_button_on_enter(e):
    run_b2['background'] = 'Black'
    run_b2['foreground'] = 'white'


def generate_button_on_leave(e):
    run_b2['background'] = 'SystemButtonFace'
    run_b2['foreground'] = 'black'


run_b2.bind("<Enter>", generate_button_on_enter)
run_b2.bind("<Leave>", generate_button_on_leave)

# Arrange Button in same grid cell
run_b2.pack(side="right")
c_button.pack(side="right", padx=10)

# Row = 9

# LinkedIn / Medium / GitHub Contacts in same grid cell using frame -- Column
# Define Frame
logo_frame = tk.Frame(app)
logo_frame.grid(row=9, padx=10, pady=5, ipadx=1, ipady=3, sticky='news', columnspan=2)


# Define a function to open webpage on click
def callback(url):
    webbrowser.open_new_tab(url)


# LinkedIn
linkedin_logo = tk.PhotoImage(file="img/linkedin_logo.png")
l_button = tkinter.Button(logo_frame, text="LinkedIn", image=linkedin_logo, borderwidth=0, cursor='hand2',
                          command=lambda: callback('https://www.linkedin.com/in/ashish568/'))

# GitHub
git_logo = tk.PhotoImage(file="img/github_logo.png")
g_button = tkinter.Button(logo_frame, text="GitHub", image=git_logo, borderwidth=0, cursor='hand2',
                          command=lambda: callback('https://github.com/ashishkr568'))

# Medium
medium_logo = tk.PhotoImage(file="img/medium_logo.png")
m_button = tkinter.Button(logo_frame, text="Medium", image=medium_logo, borderwidth=0, cursor='hand2',
                          command=lambda: callback('https://medium.com/@ashish.568'))

# Arrange Buttons in single grid
l_button.pack(side="left")
g_button.pack(side="left", ipadx=10)
m_button.pack(side="left")

# Resize Grid
app.columnconfigure(tuple(range(3)), weight=1)
app.rowconfigure(tuple(range(9)), weight=0)

# Close the App
app.mainloop()
