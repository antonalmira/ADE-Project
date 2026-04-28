ADE_PROJECT_ver3/
├── __pycache__/                 # Python cache files (auto-generated)
├── myenv/                       # Virtual environment folder
├── resources/                   # Folder with GUI icons
├── build/
├── dist/
│   ├── main.exe				  # Application
├── src/                         # Source code directory
│   ├── __pycache__/             # Python cache files (auto-generated)
│   ├── app.py                   # Main application logic
│   ├── chart_extractor.py       # Chart extraction from Excel
│   ├── document_generator.py    # Detailed document generation logic
│   ├── document_handler.py      # High-level document generation/update
│   ├── excel_handlers.py        # Excel chart handling logic
│   ├── excel_utils.py           # Excel table extraction utilities
│   ├── handlers.py              # UI event handlers
│   ├── image_utils.py           # Image cropping utilities
│   ├── list_updater.py          # List updating logic
│   ├── main.py                  # Application entry point
│   ├── performance_section.py   # Performance data handling
│   ├── preview.py               # File preview logic
│   ├── resource_rc.py           # Compiled Qt resource file
│   ├── SAMPLE.PY                # Sample Python script/ only used for testing importing of tables with copied format/ delete after 
│   ├── utils.py                 # General utility functions
│   ├── waveform_section.py      # Waveform image handling
│   ├── word_utils.py            # Word document utilities
├── templates/                   # Folder for template files
│   ├── DER-template.docx        # Word document template
├── DocuApp_ver4.ui.py
├── DocuApp_ver4.py
├── DocuApp_ver4.ui
├── DocuApp_ver5.ui
├── DocuApp_ver5.py
├── DocuApp_ver6.ui
├── DocuApp_ver6.ui.py
├── installing_dependencies.txt  # Instructions for installing dependencies
├── performancedata_testnames.json
├── README.txt                   # Project documentation
├── resource.qrc                 # Qt resource file
├── resource_rc.py               # Compiled Qt resource file
├── waveform_testnames.json      # Waveform test names configuration




********FOR THE NEXT DEVELOPER**********

To run this project, install dependencies first:
To run this project, install dependencies first:
pip install pyqt5 pyinstaller openpyxl pillow pywin32 pyautogui python-docx
install pyqt5-tools as well if GUI needs to be edited as well, alternitively, install the qt designer separately

pip install pyqt5 pyinstaller openpyxl Pillow pywin32 pyautogui python-docx

Install pyqt5-tools as well if GUI needs to be edited, alternatively, install qtdesigner separately. 

sample folder under the ADE Project main folder is used for sample inputs in the ADE app.

Update executable once finished with modifications.



**FEATURES TO DEVELOP**

├──caption formats: ensure that under performance data handling, captions would be generalized. Try to apply how the waveform caption format handling where the filename is used as the caption to performance data handling as well.
├──create a reset button for default tests (ask Mark for more info)
├──fix progress window for generating or updating: problem is it does not pop up quickly
├──implement another way to extract tables: use the process indicated in "SAMPLE.py" where you copy and paste the table along with the format (e.g, Bold, Font Size, Cell Color). When successfully implemented, change the standard requirements on User Guide ppt to follow POWI format for table. 
├──apply tm rules at the very first instance of a unit (ask DocCon for more info)
├──document properties clear title, author and subject: or put it during output (ask DocCon for more info)
├──instead of extracting charts as images, make it an editable chart: would prefer if connected to the excel file where the chart is coming from: if edited in word, excel will update automatically and vice versa. If hard to implement, export the editable chart without any link to the excel
├──try different input formats and error handling
├──DocuApp_ver6.ui is the latest ui, all other ui files (e.g, DocuApp_ver4-5.ui) are previous versions of the ui incase of fallbacks




**********Czarina Bonrostro - Intern Data Automation 2025**********

