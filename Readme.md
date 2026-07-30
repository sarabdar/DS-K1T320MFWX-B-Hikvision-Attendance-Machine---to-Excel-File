# Hikvision Attendance Excel Generator

A Python tool that fetches attendance logs from a **Hikvision DS-K1T320MFWX-B** (or compatible) device via its ISAPI and exports them into a styled Excel workbook — one sheet per employee, with daily in/out times and total hours worked.

---

## Project Structure

```
hikvision_attendence/
├── main.py            # Entry point — orchestrates the full pipeline
├── config.py          # Loads settings from .env
├── cli/               # CLI prompts & user interaction
├── hikvision/         # HTTP client & paginated event fetching
├── attendance/        # Data transformation & hour calculations
├── export/            # openpyxl Excel styling & writing
├── requirements.txt   # Python dependencies
├── .env-example       # Template for  credentials
└── .env               # Actual credentials 
```

---

## Setup

### 1. Create a Virtual Environment

```bash
python3 -m venv venv
```

Activate it:

```bash
# Linux / macOS
source venv/bin/activate

# Windows
venv\Scripts\activate
```

> You should see `(venv)` at the start of terminal prompt.

---

### 2. Install Requirements

```bash
pip install -r requirements.txt
```

This installs:
- `requests` — HTTP communication with the Hikvision device
- `pandas` — attendance data transformation
- `openpyxl` — Excel file creation and styling
- `python-dotenv` — loads credentials from `.env`

---

### 3. Configure Environment Variables

Copy the example file and fill in device credentials:

```bash
cp .env-example .env
```

Edit `.env`:

```env
DEVICE_IP="192.168.1.100"
USERNAME="admin"
PASSWORD="your_password_here"
```

| Variable    | Description                              |
|-------------|------------------------------------------|
| `DEVICE_IP` | IP address of  Hikvision device      |
| `USERNAME`  | Admin username for the device            |
| `PASSWORD`  | Admin password for the device            |

> ⚠️ **Never commit `.env` to version control.** It is already listed in `.gitignore`.

---

### 4. Run the Application

```bash
python main.py
```

The script will:
1. **Prompt you** to select the month for which to generate the report.
2. **Fetch** all attendance events from the device for that month.
3. **Transform** the raw events into a pivot table (first-in / last-out per day per employee).
4. **Export** a styled `.xlsx` Excel file to the current directory.

---

## Output

The generated Excel file is named after the selected month (e.g., `Attendance_July_2026.xlsx`) and contains:
- One row per employee per day
- Columns for **Check-In**, **Check-Out**, and **Total Hours**
- Styled formatting for easy reading

---

## Troubleshooting

| Problem | Solution |
|---|---|
| `Connection refused` or timeout | Verify `DEVICE_IP` is correct and the device is reachable on the network |
| `401 Unauthorized` | Double-check `USERNAME` and `PASSWORD` in `.env` |
| `No attendance logs found` | Ensure the device has events recorded for the selected month |
| `ModuleNotFoundError` | Make sure the virtual environment is activated and `pip install -r requirements.txt` was run |
