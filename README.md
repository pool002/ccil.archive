## CCIL Archive Downloader

### Download. Archive. Transform.

A Python-based utility to download, archive, sanitize, and transform CCIL Excel data into structured and formatted workbooks.

⚡ Quick Start

Follow these commands to get started.

```bash
# Clone the repository
git clone <repository-url>

# Navigate to the project directory
cd ccil.archive

# Create a virtual environment
python3 -m venv venv

# Activate the environment (Linux / macOS)
source venv/bin/activate

# Upgrade pip
python -m pip install --upgrade pip

# Install dependencies
pip install -r requirements.txt

# Run the application
python3 archive-downloader.py
```

Windows (Command Prompt): Replace the activation command with:

```
venv\Scripts\activate
```
-----------------------------------------------------------------
📋 Features
Download raw archive data from CCIL.
Export downloaded data into Excel (.xlsx) format.
Organize data into individual worksheets.
Sanitize source data before exporting.
Calculate aggregate totals and append a Total row.
Apply structured Excel formatting.
Preserve the original raw archive.
📁 Project Structure
```
ccil.archive/
│
├── venv/                       # Python virtual environment
│
├── archive-downloader.py       # Main application
├── beautify.py                 # Excel transformation module
├── requirements.txt            # Project dependencies
├── README.md                   # Project documentation
│
└── ...                         # Generated Excel archives
```

⚙️ Prerequisites
```
Python 3.14 (or a compatible Python version)
pip
Git (for cloning the repository)
```
Verify your installations:
```
python3 --version
pip --version
```

🛠️ Installation & Setup
1. Clone the Repository
bash
git clone <repository-url>
cd ccil.archive

2. Create a Virtual Environment
bash
python3 -m venv venv

3. Activate the Environment

Linux / macOS

bash
source venv/bin/activate


Windows (Command Prompt)

cmd
venv\Scripts\activate


The terminal prompt should display (venv) when the environment is active.

4. Install Dependencies
bash
python -m pip install --upgrade pip
pip install -r requirements.txt


Current dependency versions:

Package	Version
pandas	3.1.1
openpyxl	3.1.5
requests	2.34.2
▶️ Running the Application

Ensure the virtual environment is activated and execute:

bash
python3 archive-downloader.py


The application downloads the CCIL archive and processes the retrieved Excel data.

📊 Output & Transformation

Successful execution generates two Excel files:

1. Raw Archive
text
<user_val>_raw_archive_<timestamp>.xlsx


Contains the original downloaded archive.

2. Beautified Archive
text
beautified_<user_val>_raw_archive_<timestamp>.xlsx


Contains the processed and formatted workbook.

Beautification Process

The beautify.py module performs the following operations:

Reads the result1 worksheet.
Removes the unnecessary dfar_id column.
Maps source columns to the required output headers.
Arranges columns in the prescribed order.
Calculates aggregate totals.
Appends a Total row.
Applies Excel formatting.
Exports a separate formatted .xlsx file.
Note: The original raw archive is retained and never overwritten by the beautification process.
🔧 Troubleshooting
Issue	Suggested Solution
ModuleNotFoundError	Activate venv and reinstall dependencies.
Permission denied	Check directory and file permissions.
Sheet 'result1' not found	Verify that the downloaded workbook contains the required worksheet.
Excel transformation failed	Verify required source columns and Excel file integrity.
Dependency installation failed	Upgrade pip and retry installation.

To reinstall dependencies:

bash
pip install -r requirements.txt

📝 Important Notes
Always activate the virtual environment before running the application.
Do not manually modify files inside venv/.
Keep archive-downloader.py, beautify.py, and requirements.txt in the project directory.
Do not commit the virtual environment to Git.
Back up downloaded archive files as required.
Deactivate the Virtual Environment

When finished, execute:

bash
deactivate


CCIL Archive Downloader

Download. Archive. Transform.
