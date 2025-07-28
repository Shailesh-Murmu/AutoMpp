Automation System Pro
📖 Overview
Automation System Pro is a desktop application designed to automate a variety of repetitive office tasks. It provides a user-friendly graphical interface (GUI) to configure and execute tasks, and a separate headless script to run scheduled automations in the background without user intervention.

Key Features:
Email Automation: Send bulk emails using templates and recipient lists from Excel files.

Google Drive Integration: Download entire folders from Google Drive, automatically converting Google Docs/Sheets/Slides to standard formats.

Tracker Generation: Consolidate data from a master Excel file and a Google Forms response sheet into a formatted, color-coded tracking spreadsheet.

Google Form Updater: Dynamically update dropdown menus in a Google Form with data from an Excel file.

Scheduled Reminders: (Coming Soon) Set up automated email reminders based on tracker data.

AI-Powered Error Diagnosis: Uses the Gemini API to provide simple, human-readable explanations and solutions for technical errors.

⚙️ Setup and Installation
Follow these steps to set up the project on your local machine.

1. Prerequisites
Python 3.8 or newer

Git

2. Clone the Repository
Open your terminal or command prompt and clone the project:

git clone <your-repository-url>
cd <your-repository-folder>

3. Set Up a Virtual Environment
It is highly recommended to use a virtual environment to manage project dependencies.

# Create a virtual environment
python -m venv venv

# Activate the virtual environment
# On Windows:
venv\Scripts\activate
# On macOS/Linux:
source venv/bin/activate

4. Install Dependencies
Install all the required Python libraries using the requirements.txt file.

pip install -r requirements.txt

5. Configure Google APIs and Credentials
This application requires access to Google Drive, Sheets, and Forms.

Enable APIs: Follow the Google API Python Quickstart guide to create a new project in the Google Cloud Console.

Enable the following APIs for your project:

Google Drive API

Google Sheets API

Google Forms API

Download Credentials: After enabling the APIs, create OAuth 2.0 Client ID credentials. Download the credentials file and rename it to credentials.json.

Place the File: Move the credentials.json file into the root directory of the project.

Important: Do NOT commit credentials.json to Git. The included .gitignore file is configured to prevent this.

6. Create the Environment File (.env)
Create a file named .env in the project's root directory. This file will store your secret keys and credentials. Add the following content, replacing the placeholder values with your actual information:

# Your Gmail address for sending automated emails
AUTOMATION_SMTP_EMAIL="your-email@gmail.com"

# Your Google Account "App Password".
# Get one here: [https://myaccount.google.com/apppasswords](https://myaccount.google.com/apppasswords)
AUTOMATION_SMTP_PASSWORD="your-google-app-password"

# Your Google AI Studio API Key for error diagnosis
# Get one here: [https://aistudio.google.com/app/apikey](https://aistudio.google.com/app/apikey)
GEMINI_API_KEY="your-gemini-api-key"

🚀 How to Use the Application
1. First-Time Setup (GUI)
When you run the application for the first time, you need to configure your settings.

Run the GUI:

python your_main_script_name.py

Authenticate with Google: The first time you run a task that requires Google access (like Drive, Tracker, or Form Updater), a browser window will open asking you to log in to your Google account and grant permission. After you approve, a token.json file will be created. You only need to do this once.

Navigate to the ⚙️ Settings Tab:

Enter your SMTP Email and SMTP App Password. This is required for the Email Task.

Enter your Gemini API Key. This is optional but recommended for helpful error messages.

Customize the appearance (background, opacity, spinner) if you wish.

Click "Save Settings".

2. Using the GUI (your_main_script_name.py)
The GUI is for creating, managing, and manually running your automation tasks.

📧 Email Task: Create tasks to send bulk emails. Specify a title, subject, message body, and an Excel file containing a column of recipient emails.

📁 Drive Folder: Create tasks to download the complete contents of a Google Drive folder to your computer.

📈 Tracker: Create tasks to generate a formatted report by comparing a master list of people/tasks (from Excel) with their submissions (from a Google Form).

📝 Form Updater: Link a Tracker Task to a Google Form to automatically update its dropdown options (e.g., Location, SPOC Name) from the master Excel file.

3. Background Automation (headless.py)
The headless.py script is designed to run in the background to execute scheduled tasks (like daily email reminders) without needing the GUI to be open.

How to Set Up Autorun on Windows Startup:

Create a Shortcut:

Right-click on the headless.py file in your project folder.

Select Send to > Desktop (create shortcut).

Open the Startup Folder:

Press Win + R to open the Run dialog.

Type shell:startup and press Enter. This will open the Startup folder for your user account.

Move the Shortcut:

Drag and drop the headless.py shortcut from your Desktop into the Startup folder.

Now, every time you log in to your computer, the headless.py script will automatically start running in the background, checking for and executing any scheduled tasks.