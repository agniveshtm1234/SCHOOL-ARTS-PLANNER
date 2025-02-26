# 🎨SCHOOL-ARTS-PLANNER📃
# 📌Overview
This **School Arts Planner** is a python-based management system designed to organize and streamline school arts festivals. It helps administrators to register events, manage participants, track winners and generate certificates all in one interface.

# ⚡Features
✅ **Event Management** - Register,Update,Delete and Display event details(in EXCEL files).

✅ **Participant Registration** - Enroll students in different competitions.

✅ **Winner Tracking** - Upload and Modify winners for each events.

✅ **Automated Certificate Generation** - Create pariticpation and winner's certificates with school logo's.

✅ **Excel Reports** - Generate structured Excel files for participants, winners and events which will get automatically opened after the data's are inserted.

✅ **Database Integration** - Uses MySQL for secure data storage.

# 🛠️Requirements
**• Python 3.x**

**• MySQL**

**• Required Python Libraries**

 ```bash
 pip install mysql-connector-python openpyxl opencv-python tkinter
 ```

# 📂Folder Structure
```bash
SCHOOL-ARTS-PLANNER/
|-- cert/ *(Stores Generated Certificates)*
|   |--- participation_certificates/
|   |--- winners_certificates/
|
|-- EXCEL TEMPLATES/ *(Contaies pre-formatedd Excel templates)*
|   |--- ITEMS(Template).xlsx
|   |--- PARTICIPANTS(Template).xlsx
|   |--- WINNERS(Template).xlsx
|
|-- INFO/ *(Stores generated Excel Reports)*
|
|-- source_code/
|   |--- SCHOOLARTSPLANNER.py *(MAIN PYTHON SCRIPT)*
|
|-- template/ *(Contains Certificate templates and logo's)*
|   |--- FIRST.jpg
|   |--- SECOND.jpg
|   |--- THIRD.jpg
|   |--- participant.jpg
|   |--- logo.png
|
|
|-- README.md
```

# 🔧Future Enhancements
🔹Add a **GUI** for better user experience.

🔹Implement **Cloud Storage** for certificate management.

🔹Enhance **reporting features** with detailed analytics.




   
