### CCIL Archive Downloader

A Python-based utility to download, archive, and transform CCIL Excel data into a structured and formatted Excel workbook.

The application uses a Python virtual environment (venv) to manage dependencies and keep the project isolated from system-wide Python packages.
----------------------------------------------------------------------------------------------
Features:
Download raw archive data from CCIL.
Export downloaded data into Excel (.xlsx) format.
Organize data into individual worksheets.
Sanitize downloaded data before exporting.
Calculate aggregate totals and generate a Total row.
Preserve the original raw Excel archive.

Project Structure:
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
----------------------------------------------------------------------------------------------
Prerequisites

Ensure the following are installed:
- Python 3.14 (or a compatible Python version)
- pip
- Git (if cloning the repository)

Check your Python installation:

python3 --version

Check pip:

pip --version

Getting Started
1. Clone the repository

git clone <repository-url>

Navigate into the project directory:

cd ccil.archive

2. Create a virtual environment

Create an isolated Python environment:

python3 -m venv venv

3. Activate the virtual environment

Linux / macOS:
source venv/bin/activate


Windows (Command Prompt):

venv\Scripts\activate


Once activated, your terminal should display (venv).

4. Upgrade pip

python -m pip install --upgrade pip

5. Install dependencies

Install all required packages:

pip install -r requirements.txt


The current dependency versions are:

Package	Version
pandas	3.1.1
openpyxl	3.1.5
requests	2.34.2
Running the Application

Make sure your virtual environment is activated.
----------------------------------------------------------------------------------------------

### Execute the main application:

python3 archive-downloader.py


The application will initiate the archive download and process the retrieved data.

Output

Upon successful execution, the application generates Excel archive files.

Raw archive:
<user_val>_raw_archive_<timestamp>.xlsx

Beautified archive:
beautified_<user_val>_raw_archive_<timestamp>.xlsx


The beautification process:
Reads the result1 worksheet.
Removes the unnecessary dfar_id column.
Maps source columns to the required output headers.
Arranges columns in the prescribed order.
Calculates aggregate totals.
Appends a Total row.
Applies Excel formatting.
Exports a separate formatted .xlsx file.

The original raw archive is retained.
----------------------------------------------------------------------------------------------
Deactivating the Virtual Environment

Once execution is complete, deactivate the virtual environment:

deactivate
----------------------------------------------------------------------------------------------
Troubleshooting
Issue	Suggested solution
ModuleNotFoundError	Activate venv and install dependencies again.
Permission denied	Check directory and file permissions.
Sheet 'result1' not found	Verify that the downloaded workbook contains the required worksheet.
Excel transformation failed	Verify the required source columns and Excel file integrity.
Dependency installation failed	Upgrade pip and retry installation.

To reinstall dependencies:

bash
pip install -r requirements.txt

Important Notes
Always activate the virtual environment before running the application.
Do not manually modify files inside the venv directory.
Keep archive-downloader.py, beautify.py, and requirements.txt in the project directory.
Do not commit the virtual environment to Git.
Keep downloaded archive files backed up as required.
Quick Start

For users who want to get started quickly:

bash
# Clone repository
git clone <repository-url>

# Navigate to project
cd ccil.archive

# Create virtual environment
python3 -m venv venv

# Activate virtual environment
source venv/bin/activate

# Install dependencies
pip install -r requirements.txt

# Run application
python3 archive-downloader.py


CCIL Archive Downloader

Download. Archive. Transform.
