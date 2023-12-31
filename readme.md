
# <img src="icon.ico" alt="drawing" width="30"/> Automatic Document Generator from Template

<img src="images/preview.png" alt="drawing" width="500"/>

This application is useful for creating certificates, diplomas, or any other type of document that requires a **large number of documents** with the same format but **different data**.

Facilitates  automatic creation of Word documents (.docx) or PDF from a predefined template. The template should contain placeholder markers, such as `[name]`, `[id]`, which will be replaced with specific values.

## How to use

1. **Document Template:**
   - Create a Word document with placeholder markers that match the column names in your Excel file. For example: `[name]`, `[id]`.
  
   ![Naming Convention](images/screenShot.png)

2. **Excel File:**
   - Prepare an Excel file with data. Ensure that the columns have the same names as the markers in the template.

    ![Excel File](images/screenShot2.png)

3. **Open .exe file:**
   - Open the .exe file and follow the instructions.
   ![Excel File](images/screenShot3.png)

4. **Fields:**
   - Fill in the fields with the path of the doc template and the Excel file.
   - *File name* is optional, it will be the prefix of the documents.
   - *Select the column*: Select the column that will be used to name the documents.
   - File name and select the column will be concatenated to name the documents.
    ![Excel File](images/preview_completed.png)
   - You can generate PDF files by checking the -**PDF?**- box too.
  ![Excel File](images/pdf.png)

5. **Done:**
   - Automatically the script will create the documents with the data of the Excel file.
   ![Excel File](images/screenShot5.png)
6. **Docx result:** :white_check_mark:
   ![Excel File](images/screenShot6.png)
  
## If you want to run the script directly

### Requirements

Make sure you have the necessary libraries installed:

```bash
pip install pandas python-docx
```

### Run

```bash
python main.py
```

### Generate .exe file

```bash
pip install cx_Freeze
```

```bash
python setup.py build
```

## Credits

This script was created by me, [@Isai_hernandez]
