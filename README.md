# Automating List of Documents

This PowerShell script helps manage collections of documents by creating simple indexes from the file names. It is designed for situations where the number of documents is too small to justify the cost of full e‑discovery tools, but a list of documents would assist with locating relevant files to review .

The script runs locally on your computer, so confidentiality of client documents is maintained.

---

## What You Need to Edit Before Running

Open the script file (`IndexFiles.ps1`) in a text editor (such as Notepad or Visual Studio Code). At the top of the script you will see three settings:

- **RootPath**  
  This is the folder where your documents are stored.  
  Example:  
  $RootPath = "C:\Desktop\File1"

 - **OutDir**
This is the folder where the script will save the CSV index files.
Example:
$OutDir = "C:\Desktop\File1\Admin"

- ** Keywords**
Inside the script there is a section that defines how “keyword” files are identified. You can change the words to suit your matter.
Examples:
1. $keywordsIndex = $rows | Where-Object { $_.'Document name' -match '(?i)purchase[ _\-]?order' } : this will return any files with the phrase "purchase order" in the file name.

2. $keywordsIndex = $rows | Where-Object { $_.'Document name' -match '(?i)(invoice|invoices|purchase order|contract|agreement)'}
- In this example, the script will look for document names containing any of the following phrases:
- “invoice” or “invoices”
- “purchase order”
- “contract”
- “agreement”
You can add or remove terms depending on the types of documents you expect.

There is an example within the script to search for the word "transcript"

--

What the Script Produces
When run, the script will create:
- A CSV file listing all documents in the folder.
- A CSV file listing documents whose names contain your chosen keywords.
- A CSV file listing transcripts with date‑time stamps in the filename.
- A text file showing the total number of files found.

Each CSV includes:
- Document name
- Date (if found in the filename)
- Reference number (if found in the filename)
- Source file path

--

## How to Run the Script
- Open PowerShell
- Press Win + R, type powershell, and press Enter.
- Go to the folder where the script is saved
Example:
cd "C:\Users\YourName\Documents"
- Run the script
.\IndexFiles.ps1
- Check the output
- The script will display progress messages in the PowerShell window.
- The CSV and text files will be saved in the folder you set as $OutDir.

License
This project is licensed under the MIT License, allowing free use, modification, and distribution.

