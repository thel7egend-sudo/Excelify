\# Safe Edit Scope



AI agents must restrict edits to the requested feature area.



---



\# Safe Files To Modify



UI improvements:



editor\_page.py  

doc\_editor\_page.py  

top\_chrome.py  

home\_page.py  



Feature additions:



main\_window.py  



---



\# Restricted Files



Changes allowed ONLY with explicit approval:



models/table\_model.py  

views/table\_view.py  

document.py  

storage.py  



---



\# Forbidden Without Approval



\- rewriting data structures

\- modifying undo/redo engine

\- changing swap system logic

\- modifying persistent file format



---



\# Rule



If a requested feature requires touching restricted files,

the AI must ask before editing them.

