#	Spreadsheet mail merge project

##	Use

The assignment file `access.log` must be in the same directory as the script, then run:


##	Run instructions

###	Test Run

Test the program from the terminal using:

	$ python main.py

The GUI will start up.
The default input files already populate all the input fields.
Press `SPACE` to hit the "Send emails" button.


###	Send emails

Open Microsoft Outlook and login to an account.

Run the program.

	$ python main.py

- Uncheck the button `Test: Don't actually send the email`.
- Create your own email template.
- Use your own CSV file.

##	Directory contents

Documentation

-	README.md   # This file
-	todo.md     # A list of to-do's for the project

Program files

-	main.py       # Entry point for script
-	gui.py        # Runs GUI and collects document names
-	emailer.py    # Emailer tranforms the documents into emails, which are then sent
-	inspector.py  # Inspect inputs and outputs of functions
-	style.css     # Default CSS file

Test files for running the code

-	test_attachment.pdf
-	test_csv.csv
-	test_template.txt



##	Project plan

1.	Begin with
	1.	CSV containing student and teacher data (extracted via from a school database).
	1.	Email template (markdown)
2.	Convert to a mass email where each email is formatted by programmatically slicing the CSV into a table depending on the email recipient.
	-	Consider an excel pivot table of student records organized by professor and class.
	-	Each email has subsection of that pivot table.

The above actions were previously performed using the Microsoft Word "Mail Merge" function, however the tables had to manually be created by staff.
This process had to be completed more than once a semester, each time requiring fifty hours of tedious labor.

###	Pseudocode

	Level 0 Outline
		Use GUI to collect information, such as the list of emails, email template file, etc.
		Convert data files into HTML files containing a formatted email
		Send each HTML file as an email via Microsoft Outlook

	Level 1 Outline
		Open GUI to collect information
			Subject line
			Email template file
			Email template CSS file
			CSV file
			Attachment file
		Open clicking SEND, begin processing the files to send emails
			Read the CSV file and organize it into a dictionary by email recipient (faculty email)
			Read the CSS file
			Use the pandoc to format the email into HTML, guided by the CSS file
				Formatting includes inserting the faculty name and the table of students organized by class.
			Send the HTML file by hooking into Outlook via the win32com library.

###	Modules used

-	csv: Read CSVs.
-	os: Install libraries for the user if they aren't already installed.
-	pandoc:  Convert from Markdown to HTML.
-	pathlib: Sensible filepath syntax.
-	pypandoc: Python API to Pandoc.
-	tkinter: GUI framework.
-	win32: Hook into Microsoft Outlook to send emails.

##	Goals

- [x] Use a CSV file to automatically send emails.
- [x] Package the program into a GUI.

##	Citations

1. CSV Library. Python.org. https://docs.python.org/3/library/csv.html
1. PyPandoc. https://pypi.org/project/pypandoc/
1. Tkinter. File Browser. TutorialsPoint.com. https://www.tutorialspoint.com/creating-a-browse-button-with-tkinter
1. Tkinter. RealPython.com. https://realpython.com/python-gui-tkinter/
1. Win32. Microsoft Outlook. https://win32com.goermezer.de/microsoft/ms-office/send-email-with-outlook-and-python.html
