\# Excelify Architecture



\## 1. Overview



Excelify is a desktop data-entry productivity application built with \*\*Python + PySide6\*\*.



The goal is \*\*not to replace Excel\*\*, but to provide a \*\*faster and smoother data entry environment\*\* for freelancers and professionals who later export their work to Excel. :contentReference\[oaicite:1]{index=1}



Excelify focuses on:



• Fast keyboard-first data entry  

• Powerful editing workflows  

• Clean UI  

• Export compatibility with Excel  



The system uses a \*\*modular architecture\*\* with clear separation between:



• UI  

• Data models  

• Storage  

• Application controller  



---



\# 2. High Level Architecture





app.py

↓

MainWindow (Application Controller)

↓

UI Layer

├── TopChrome

├── HomePage

├── EditorPage (Grid editor)

└── DocEditorPage (Document editor)



↓

Model Layer

├── Document

├── Sheet

└── TableModel



↓

View Layer

└── TableView



↓

Storage Layer

└── storage.py





The system roughly follows a \*\*Model-View-Controller inspired architecture\*\*.





Model → Document / Sheet / TableModel

View → TableView / UI widgets

Controller → MainWindow / EditorPage





---



\# 3. Application Entry Point



\### app.py



This is the program entry point.



Responsibilities:



• Create Qt application  

• Create the main window  

• Start event loop  



Example:



```python

app = QApplication(sys.argv)

window = MainWindow()

window.show()

app.exec()



app



4\. Main Application Controller

MainWindow



MainWindow is the central controller of the application.



Responsibilities:



• Switching between pages

• Managing documents

• Import/export

• Saving app state

• Dark mode control



It manages two primary UI states:



Home Mode

TopChrome

HomePage

Editor Mode

TopChrome

EditorPage OR DocEditorPage



Editor type depends on document type:



grid → EditorPage

doc → DocEditorPage



main\_window



5\. UI Components

TopChrome



TopChrome is the top navigation bar.



Features:



• Home button

• Document search

• Search results dropdown



Signals:



home\_clicked

search\_selected



It can switch between:



home mode

editor mode



top\_chrome



6\. Home Page



HomePage is the document manager screen.



Responsibilities:



• Display all documents

• Create/open documents

• Import Excel files

• Toggle dark mode

• Trigger document opening



Main signals:



open\_document\_requested

open\_import\_requested

save\_requested

toggle\_dark\_mode\_requested

7\. Document System



The application revolves around the Document model.



A document can be two types:



grid  → spreadsheet-style data

doc   → text document



Structure:



Document

&nbsp;├── name

&nbsp;├── type

&nbsp;├── sheets\[]

&nbsp;├── active\_sheet\_index

&nbsp;└── content (doc type only)



Each document contains one or more sheets.



document



8\. Sheet Model



Each spreadsheet document contains Sheet objects.



Sheet stores the actual data.



Structure:



Sheet

&nbsp;├── name

&nbsp;├── cells {(row,col): value}

&nbsp;├── row\_heights

&nbsp;└── col\_widths



Cells are stored in a sparse dictionary:



cells = {

&nbsp; (row, col): value

}



This design is efficient for large spreadsheets because empty cells are not stored.



document



9\. Grid Editor (EditorPage)



EditorPage is the main spreadsheet editor.



Responsibilities:



• Spreadsheet editing

• Ribbon tools

• Zoom Box editing

• Swap operations

• Sheet management

• Undo/Redo



Core UI structure:



EditorPage

&nbsp;├── Ribbon

&nbsp;│   ├── Swap Cell

&nbsp;│   ├── Swap Row

&nbsp;│   ├── Swap Column

&nbsp;│   ├── Zoom Box

&nbsp;│   ├── Undo

&nbsp;│   ├── Redo

&nbsp;│   └── Export to Excel

&nbsp;│

&nbsp;├── TableView

&nbsp;│

&nbsp;├── Zoom Box

&nbsp;│

&nbsp;└── Sheet Bar



editor\_page



10\. Table Model

TableModel



TableModel is the core data model used by the spreadsheet view.



It inherits:



QAbstractTableModel



Responsibilities:



• Providing cell data to the view

• Handling edits

• Tracking history

• Undo/redo system

• Swap operations

• Batch operations



Features:



Undo / Redo



Implemented with:



\_undo\_stack

\_redo\_stack



Supports:



compound actions

macro operations

per-sheet history



table\_model



11\. Table View

TableView



TableView is the interactive spreadsheet UI.



Based on:



QTableView



Responsibilities:



• Rendering spreadsheet grid

• Managing selection

• Handling drag swap

• Context menu operations

• Clipboard operations



Features:



Clipboard Support

Copy

Cut

Paste

Delete

Transform Tools

Remove spaces

Uppercase

Swap Drag System

drag\_swap\_requested

block\_swap\_requested



table\_view



12\. Zoom Box Editing System



Zoom Box is a special editing mode designed for fast data entry.



When enabled:



• Table editing is disabled

• Editing occurs inside a large text box



Features:



Segment splitting

Marker-based splitting

Arrow navigation

Row/column auto advance



Segment markers allow splitting text into multiple cells.



Example:



123|456|789



becomes



A1 = 123

A2 = 456

A3 = 789



editor\_page



13\. Document Editor

DocEditorPage



DocEditorPage is used for text documents.



Features:



• Multi-page layout

• Automatic pagination

• Rich text editing



Pages are dynamically generated based on text height.



PAGE\_WIDTH

PAGE\_HEIGHT

PAGE\_MARGIN



Text is automatically split across pages.



doc\_editor\_page



14\. Storage System



Storage is handled through a simple JSON persistence layer.



Module:



storage.py



Functions:



save\_state()

load\_state()



Data is stored in:



data/app\_state.json



Each document is serialized using:



Document.to\_dict()



and restored using:



Document.from\_dict()



storage



15\. Import / Export System

Excel Import



Excel files are imported using:



openpyxl



Process:



Excel → Document → Sheet → Cells



Each Excel worksheet becomes a Sheet.



main\_window



Excel Export



Export process:



Document

&nbsp;↓

Sheets

&nbsp;↓

Cells

&nbsp;↓

openpyxl Workbook

&nbsp;↓

.xlsx file



Only populated cells are written.



main\_window



16\. Dark Mode System



Excelify uses two dark mode layers.



Application Dark Mode



Controls:



dialogs

menus

global UI

Grid Dark Mode



Controls:



table

editor ribbon

sheet bar

zoom box



Dark mode is toggled via:



toggle\_dark\_mode()



main\_window



17\. Signals \& Data Flow



Important signal flow:



EditorPage → document\_changed → MainWindow → save\_state

TableModel → save\_requested → MainWindow → save\_state

HomePage → open\_document\_requested → MainWindow → EditorPage

EditorPage → export\_requested → MainWindow → export\_excel

18\. Current Design Philosophy



Excelify focuses on:



fast data entry

keyboard workflows

simple data models

Excel compatibility



It intentionally avoids:



complex spreadsheet formulas

Excel-like feature bloat

heavy computation



Instead the workflow is:



Excelify → fast entry → export → Excel for heavy work



README



19\. Future Architecture Extensions



Planned architectural expansions:



OCR Pipeline

Image → OCR → Grid → Excelify

AI Data Cleaning

Cell → AI cleanup → structured output

Plugin System

user tools

transform scripts

custom workflows

Licensing System

offline trial

license verification

update validation

20\. Codebase Structure

excelify/

│

├── app.py

├── main\_window.py

│

├── document.py

├── storage.py

│

├── models/

│   └── table\_model.py

│

├── views/

│   └── table\_view.py

│

├── editor\_page.py

├── doc\_editor\_page.py

│

├── top\_chrome.py

├── home\_page.py

│

├── data/

│   └── app\_state.json

│

└── README.md

21\. Key Architectural Principles



Excelify follows these principles:



1\. Separation of concerns

UI

Model

Storage

2\. Sparse data storage



Efficient large spreadsheet handling.



3\. Signal-driven UI



Qt signals propagate state updates.



4\. Stateless Views



Views never store data permanently.



5\. Explicit document model



Everything revolves around:



Document → Sheet → Cells

End of Architecture



---



If you want, next I can also write a \*\*`CODEX\_ARCHITECTURE\_GUIDE.md`\*\* which is a \*\*special architecture doc designed specifically for AI coding agents (Codex / Cursor / Claude Code)\*\* so they \*\*never break Excelify again while editing the code\*\*.  



That one is \*\*extremely powerful\*\* and will save you \*\*months of debugging later.\*\*

