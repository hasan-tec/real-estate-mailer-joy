# 🏠 Real Estate Mailer App - Setup Guide

Welcome! Follow this step-by-step guide to set up your computer to run the **Trifold Real Estate Mailer Generator**. This guide is designed for Windows users.

---

## 📋 Prerequisites checklist

We will install three main tools:
1.  **Git** (To download and manage the code)
2.  **Python** (The programming language the app is built with)
3.  **VS Code** (A good editor to view/run code)

---

## 🐙 Step 1: Install Git

1.  Go to the official Git download page: [git-scm.com/download/win](https://git-scm.com/download/win)
2.  Click on **"Click here to download"** to get the latest version.
3.  Run the installer.
4.  **Click "Next"** through all the options (the default settings are perfect for what we need).
5.  Wait for the installation to finish.

### ⏳ Verify Git Installation
1.  Open Command Prompt (Start > type `cmd` > Enter).
2.  Type `git --version` and press Enter.
3.  You should see something like `git version 2.x.x`.

---

## 🚀 Step 2: Install Python

1.  Go to the official Python download page: [python.org/downloads](https://www.python.org/downloads/)
2.  Click the big yellow button **Download Python 3.x.x**.
3.  **⚠️ IMPORTANT:** When the installer opens, you **MUST** check the box that says **"Add Python to PATH"** at the bottom.
    > ![Add Python to Path](https://miro.medium.com/v2/resize:fit:1358/1*3_M5jM9X-XhYcZ2jYz0wYg.png)
    *(Make sure this box is checked!)*
4.  Click **"Install Now"** and wait for it to finish.

### ⏳ Verify Python Installation
1.  Open a **NEW** Command Prompt window (close the old one to refresh).
2.  Type `python --version` and press Enter.
3.  You should see something like `Python 3.12.x`.

---

## 💻 Step 3: Install VS Code (Optional but Recommended)

1.  Go to [code.visualstudio.com](https://code.visualstudio.com/).
2.  Download the **Windows** installer.
3.  Run the installer and click "Next" through the options. Default settings are fine.
4.  Once installed, open VS Code.

---

## 📥 Step 4: Clone the Repository

This downloads the code from the internet to your computer.

1.  Create a folder on your Desktop (or anywhere you like) where you want to keep the project.
2.  Open that folder, right-click in the empty space, and select **"Open Git Bash Here"** (or just use Command Prompt and `cd` to that folder).
3.  Type the following command and press Enter:
    ```bash
    git clone https://github.com/hasan-tec/real-estate-mailer-joy.git
    ```
4.  A new folder named `real-estate-mailer-joy` will appear. This is the project!

---

## 🛠️ Step 5: Project Setup

### 1. Open the Project
1.  Open VS Code.
2.  Go to **File > Open Folder**.
3.  Select the `real-estate-mailer-joy` folder you just cloned.

### 2. Configure the "Virtual Environment"
This creates a robust, isolated space for the app's libraries so they don't interfere with other things on your computer.

1.  In VS Code, open the **Terminal** by pressing \`Ctrl + \`\` (the backtick key, next to 1).
2.  Type the following command and press Enter:
    ```powershell
    python -m venv env
    ```
    *(This creates a folder named `env` which will hold our libraries)*

3.  **Activate** the environment.
    *   **Windows (Command Prompt):** `env\Scripts\activate`
    *   **Windows (PowerShell):** `.\env\Scripts\Activate.ps1`
    *(You should see `(env)` appear at the start of your command line)*

### 3. Install Dependencies
Now we install the specific tools the app needs (like PDF generators, map tools, etc).

1.  Make sure you see `(env)` in your terminal.
2.  Run this command:
    ```powershell
    pip install -r requirements.txt
    ```
    *(Watch the progress bars as it downloads everything!)*

---

## 🗺️ Step 6: Mapbox API Key Setup

You need a key from Mapbox so the app can generate maps for the properties.

1.  **Sign Up**: Go to [mapbox.com](https://www.mapbox.com/) and create a free account.
2.  **Get Token**: Once logged in, go to your [Account Dashboard](https://account.mapbox.com/).
3.  Arguments:
    - Look for the section **"Access tokens"**.
    - You will see a "Default public token". Click the **Clipboard icon** to copy it.
    - It should start with `pk.eyJ...`
4.  **Configure App**:
    - Go back to VS Code.
    - Find the file named **`.env.example`**.
    - Rename it to **`.env`** (Remove .example).
    - Open the file and find `MAPBOX_TOKEN=`.
    - Paste your token there:
      ```
      MAPBOX_TOKEN=pk.eyJ1Ijoi...
      ```
    - Save the file (Ctrl + S).

---

## 🏃 Step 7: Running the App

You have two ways to run the app.

### Option A: The "One-Click" Method (Recommended)
We have prepared a file named `Run_Mailer_Trifold.bat`.

1.  Go to your folder in File Explorer.
2.  Double-click **`Run_Mailer_Trifold.bat`**.
3.  A black window will appear (this is the app loading), followed by the graphical interface.
    > **Note:** Do not close the black window while using the app!

### Option B: Via VS Code Terminal
1.  Ensure your terminal shows `(env)`.
2.  Type:
    ```powershell
    python mailer_app_trifold.py
    ```

---

## 🧩 Troubleshooting & Diagrams

### App Workflow
```mermaid
graph TD
    A[Start] --> B{Is Git Installed?};
    B -- No --> C[Install Git];
    B -- Yes --> D{Is Python Installed?};
    D -- No --> E[Install Python + Add to PATH];
    D -- Yes --> F[Clone Repo with 'git clone'];
    F --> G{Is Virtual Env (env) created?};
    G -- No --> H[Run: python -m venv env];
    G -- Yes --> I{Are requirements installed?};
    I -- No --> J[Activate env & Run: pip install -r requirements.txt];
    I -- Yes --> K{Is Mapbox Token in .env?};
    K -- No --> L[Get Key from Mapbox & Save to .env];
    K -- Yes --> M[Double Click .bat File];
    M --> N[🚀 App Launches!];
    C --> B;
    E --> D;
    H --> G;
    J --> I;
    L --> K;
```

### Common Issues

**"Python is not recognized..."**
- **Fix:** Reinstall Python and ensure you check **"Add Python to PATH"**.

**"Module not found..."**
- **Fix:** You likely didn't activate the environment or install requirements.
  1. Open Terminal.
  2. `.\env\Scripts\activate`
  3. `pip install -r requirements.txt`

**"ERROR: WEASYPRINT / GTK DEPENDENCIES MISSING"**
- **Issue:** The PDF generator (WeasyPrint) sometimes needs extra system files called GTK.
- **Fix:**
  1. Download the **GTK3 Installer for Windows** from [here](https://github.com/tschoonj/GTK-for-Windows-Runtime-Environment-Installer/releases).
  2. Run the installer and select all defaults.
  3. Restart your computer (required to update PATH).
  4. Try running the app again.

**Mapbox/Geocoding Errors**
- **Fix:** Ensure your `.env` file contains your valid Mapbox API token.
  ```
  MAPBOX_TOKEN=pk.eyJ1...
  ```

---

**That's it! You're ready to generate trifold mailers! 🏠✉️**
