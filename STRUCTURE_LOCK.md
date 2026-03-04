\# Excelify Structure Lock



The following systems are considered \*\*stable architecture\*\*.



AI agents must NOT rewrite or redesign them.



Allowed actions:

\- bug fixes

\- small targeted improvements



Forbidden actions:

\- structural redesign

\- replacing systems

\- major refactoring



---



\# Locked Systems



\## TableModel



Critical logic:



\- undo / redo system

\- compound actions

\- history per sheet

\- swap operations



File:

models/table\_model.py



This system is stable and must not be rewritten.



---



\## TableView



Handles:



\- selection

\- swap drag system

\- clipboard operations

\- ghost drag system



File:

views/table\_view.py



This system must remain stable.



---



\## Document Model



Core data structure:



Document → Sheet → Cells



File:

document.py



Changing this breaks storage compatibility.



---



\## Storage System



File:

storage.py



Defines app persistence format.



Changing it risks data corruption.



---



\# Protected Principle



Excelify uses \*\*sparse cell storage\*\*:



cells = {(row, col): value}



This must never be replaced with a dense matrix.

