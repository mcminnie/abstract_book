import pandas as pd
from pylatex import Document, Section, Command, NoEscape, NewPage, Package
import re

input_file_path = 'example_input.csv'

# List of sessions, dates, and chairs
session_details = [
    {
    "session": "Topic 1",
    "date": "Thursday 30th November, 09:00am",
    "chairs": ["Chair 1", "Chair 2", "Chair 3"]
    },
    {
    "session": "Topic 2",
    "date": "Thursday 30th November, 2:00pm",
    "chairs": ["Chair 1", "Chair 2", "Chair 3"]
    },
    {
    "session": "Topic 3",
    "date": "Friday 1st December, 9:30am",
    "chairs": ["Chair 1", "Chair 2", "Chair 3"]
    },
    { "session": "Topic 4",
    "date": "Friday 1st December, 2:00pm",
    "chairs": ["Chair 1", "Chair 2", "Chair 3"]
    }
]

session_names_list = [session["session"] for session in session_details]

# Import the csv/excel file 
# df = pd.read_excel(input_file_path)
df = pd.read_csv(input_file_path)
df = df.sort_values(by='Programme Order')

# Create separate DataFrames for each topic and store in a dictionary
session_dataframes = {}
for session in session_names_list:
    session_dataframes[session] = df[df['Session'] == session]


############## Create dividers for each session with session details on ##############

# Function to sanitize session names for filenames
def sanitize_filename(name):
    # Replace spaces with underscores
    name = name.replace(' ', '_')
    # Remove or replace special characters
    name = re.sub(r'[\\/*?:&"<>|]', '', name)
    return name


title_page_list = [] # Create a list to store the names of the files

#Create a LaTeX document for each session title page 
for session in session_details:
    session_page = Document(documentclass='report')

    # Create the LaTeX code for the session page
    latex_code = (
        r'\newpage'
        r'\thispagestyle{empty}'
        r'\vspace*{\fill}'
        r'\begin{center}'
        rf'{{\Huge \bfseries {session["session"]}}}\\[0.5cm]'  # Session title
        rf'{{\LARGE {session["date"]}}}\\[1cm]'              # Session date
        rf'{{\large \textbf{{Chairs:}} {", ".join(session["chairs"])}.\\[1cm]}}'  # Chairs
        r'\end{center}'
        r'\vspace*{\fill}'
        r'\newpage'
    )

    # Add the LaTeX code to the document
    session_page.append(NoEscape(latex_code))

    # Write the LaTeX code to a .tex file with the cleaned filename (no special characters)
    filename = f"{sanitize_filename(session['session'])}.tex"

    title_page_list.append(filename) # adds name of each file to a list so it can be input to the final document later

    with open(filename, 'w') as file:
        file.write(latex_code)


# Custom margins using geometry_options in Document - allows more words on a page
geometry_options = {
    "margin": "1.3in",  # Example: 1 inch margins on all sides
    "headheight": "20pt",
    "headsep": "10pt",
    "footskip": "25pt"
}

doc = Document(documentclass='report', geometry_options=geometry_options)

# Add required packages and custom commands
doc.preamble.append(NoEscape(r'\setlength{\parindent}{0pt}')) #stops the first line of each abstract from indenting which looks silly. 
doc.packages.append(Package('titling'))
doc.preamble.append(NoEscape(r'\newcommand{\subtitle}[1]{'
                             r'\posttitle{'
                             r'\par\end{center}'
                             r'\begin{center}\large#1\end{center}'
                             r'\vskip0.5em}}'))
doc.packages.append(Package('titlesec'))
doc.preamble.append(NoEscape(r'\titleformat{\paragraph} '
                             r'{\normalfont\normalsize\bfseries}{\theparagraph}{1em}{}'))
doc.preamble.append(NoEscape(r'\titlespacing*{\paragraph}'
                             r'{0pt}{3.25ex plus 1ex minus .2ex}{1.5ex plus .2ex}'))
doc.packages.append(Package('graphicx'))

# Title Page content
doc.append(NoEscape(r'\input{title_page.tex}'))
doc.append(NewPage())

# Function to add content to the document
def add_content_to_document(row):
    with doc.create(Section(row['Title'], numbering=False)):
        doc.append(NoEscape(r'\\'))  # New line
        doc.append(row['Author(s)'])
        doc.append(NoEscape(r'\\'))  # New line
        doc.append(Command('textit', row['Affiliation(s)']))
        doc.append(NoEscape(r'\newline'))
        doc.append(NoEscape(r'\hrule'))
        doc.append(NoEscape(r'\vspace{0.7cm}')) 
        doc.append(row['Abstract'])
        doc.append(NewPage())

# #  Add title, abstract, author, affiliations content

# If you want all the abstracts in one big block with no separation pages
# for _, row in df.iterrows(): 
#     add_content_to_document(row)

# For if you want to split by section
first_session = session_names_list[0]
first_dataframe = session_dataframes.get(first_session)
doc.append(NoEscape(r'\newpage'))
doc.append(NoEscape(rf'\input{{{title_page_list[0]}}}'))
for _, row in first_dataframe.iterrows():
    add_content_to_document(row)

second_session = session_names_list[1]
second_dataframe = session_dataframes.get(second_session)
doc.append(NoEscape(r'\newpage'))
doc.append(NoEscape(rf'\input{{{title_page_list[1]}}}'))
for _, row in second_dataframe.iterrows():
    add_content_to_document(row)

third_session = session_names_list[2]
third_dataframe = session_dataframes.get(third_session)
doc.append(NoEscape(r'\newpage'))
doc.append(NoEscape(rf'\input{{{title_page_list[2]}}}'))
for _, row in third_dataframe.iterrows():
    add_content_to_document(row)

fourth_session = session_names_list[3]
fourth_dataframe = session_dataframes.get(fourth_session)
doc.append(NoEscape(r'\newpage'))
doc.append(NoEscape(rf'\input{{{title_page_list[3]}}}'))
for _, row in fourth_dataframe.iterrows():
    add_content_to_document(row)

# Generate final PDF
doc.generate_pdf('book_of_abstracts', clean_tex=False)