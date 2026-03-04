\# AI Change Protocol



Any AI agent modifying this repository must follow this protocol.



\## Step 1 — Identify the system



The AI must first determine which subsystem is being modified:



\- UI Layer

\- TableModel

\- TableView

\- Document Model

\- Storage System

\- Import/Export System



The AI must clearly state which subsystem it is editing.



---



\## Step 2 — Explain the change



Before writing code, the AI must explain:



1\. What problem it is solving

2\. Which files will change

3\. Why the change is safe

4\. Why no other systems are affected



---



\## Step 3 — Minimal edits only



The AI must modify \*\*only the necessary lines\*\*.



Forbidden actions:



\- rewriting entire files

\- refactoring unrelated code

\- changing architecture



---



\## Step 4 — Respect architecture.md



All changes must follow the architecture defined in:



architecture.md



If the requested change violates architecture, the AI must refuse.



---



\## Step 5 — No hidden behavior changes



If behavior changes, the AI must explicitly state:



\- what changes

\- why

\- which workflows are affected

