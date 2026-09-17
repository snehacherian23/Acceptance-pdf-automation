# Acceptance PDF Automation

An internal web application that automates two connected workflows for Omantel
Acceptance Test documents:

- **Workflow 1 — Generate Acceptance Test Documents**: reads one or more Excel
  sheets and generates one Acceptance Test Word document per account,
  matching the official Omantel template.
- **Workflow 2 — Prepare Final Signed Documents**: takes the signed versions of
  those documents back, converts them to PDF, and appends the Consolidated
  Sheet(s) to produce the final deliverable.

---

## Requirements

- **Python 3.10+**
- **Microsoft Word** (required for Workflow 2 — used to convert Word documents to PDF)
- The Python packages listed in `requirements.txt`

---

## First-time setup

This is a one-time step (per computer). Open Command Prompt in the project
folder and run:

```
python -m venv venv
venv\Scripts\pip install -r requirements.txt
```

After this, use `Start.bat` to run the application from now on — no need to
repeat these commands.

---

## Launching the Application

### How to start it

Double-click **`Start.bat`** in the project folder.

That's it — no need to open VS Code, open a terminal, or type any commands.

### What the launcher does automatically

When you double-click `Start.bat`, it will:

1. Locate the project folder automatically (it works no matter where the
   folder has been placed or copied to).
2. Check whether the application is already running, and if so, just open
   your browser to it instead of starting a second copy.
3. Check that the virtual environment (`venv`) has been set up, and give
   clear instructions if it hasn't.
4. Activate the virtual environment.
5. Verify Python and all required packages are available.
6. Start the application server.
7. Automatically open your default web browser to
   **http://127.0.0.1:5000** once the server is ready.

A Command Prompt window will stay open the whole time the application is
running — this is expected. It shows the server's activity log, and if
anything goes wrong, the error will be printed there instead of the window
just closing.

### How to stop it

**Simply close the Command Prompt window** (click the X, or click inside it
and press `Ctrl+C`). The server runs directly inside that window, so closing
it stops the application immediately. There is no separate "shut down" step.

A `Stop.bat` file is also included as a fallback safety net, for the unusual
case where a server process is somehow still running after the window was
closed (for example, if `Start.bat` was accidentally run more than once).
Double-clicking `Stop.bat` will find and stop anything still using the
application's port. In normal day-to-day use you shouldn't need it.

### Troubleshooting

| Problem | What to do |
|---|---|
| Window shows "virtual environment was not found" | Run the [First-time setup](#first-time-setup) commands above, then try `Start.bat` again. |
| Window shows "required Python packages are missing" | Run `venv\Scripts\pip install -r requirements.txt`, then try `Start.bat` again. |
| Browser doesn't open automatically after ~30 seconds | Open a browser yourself and go to `http://127.0.0.1:5000` — the server may just still be starting up. |
| "Port 5000 is already in use" / app seems to already be running | This is expected if you already have it open elsewhere — `Start.bat` will open your browser to the existing instance. If that's not expected, run `Stop.bat` first, then `Start.bat` again. |
| The window closed immediately with no message visible | This shouldn't happen — every error path in `Start.bat` pauses and waits for a key press before closing. If you see this, please report it along with what you were doing right before it happened. |

---

## Project structure

```
acceptance_pdf_app/
├── Start.bat              # Double-click this to run the application
├── Stop.bat                # Fallback cleanup utility (rarely needed)
├── main.py                  # Application entry point
├── config.py                 # Configuration
├── requirements.txt           # Python dependencies
├── app/
│   ├── routes/                # Flask routes (auth, dashboard, processing)
│   ├── services/               # Excel reading, document generation, PDF conversion/merging
│   ├── models/                  # Database models
│   ├── templates/                # HTML templates
│   ├── static/                    # CSS/JS
│   └── assets/icons/               # Customer Experience icon images
├── uploads/                    # Runtime: uploaded files (not committed)
├── processed/                   # Runtime: generated output (not committed)
└── instance/                     # Runtime: SQLite database (not committed)
```
