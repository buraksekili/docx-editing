# docx-editing

Required installations.
```
pip install openpyxl
pip install python-docx
```

Automation tool for creating name badges for every attendee of 'Sabanci Bilisim Gunleri 2019' event.
It parses a template name badge (in .txt form) to generate every attendees badges. 
- It changes Name, Surname and Company fragments in copy of template name badge based on the attendee's information
- Then, new copy is saved into corresponding directory as katilimcilar(which means attendees in Turkish) or konusmacilar(which means speakers in Turkish)
