# Statistical Process Capability & Optimization Tool

## Quick Start Guide

### For Windows Users

**First Time Setup:**
1. Install Python 3.10+ from https://python.org
   > ⚠️ Check "Add Python to PATH" during installation!
2. Double-click `Setup_Windows.bat`
3. Wait for installation to complete
4. Browser opens automatically at http://localhost:5180

**After Setup:**
- Just double-click `start.bat` to run

---

### For Mac/Linux Users

**First Time Setup:**
```bash
cd /path/to/py
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
streamlit run "import streamlit as st.py" --server.port 5180
```

**After Setup:**
- Double-click `start.sh` or run the commands above

---

## Files Included

| File | Purpose |
|------|---------|
| `import streamlit as st.py` | Main application |
| `requirements.txt` | Python dependencies |
| `Setup_Windows.bat` | Windows first-time setup |
| `start.bat` | Windows quick launcher |
| `start.sh` | Mac/Linux quick launcher |
| `Setup_Guide.pdf` | Printable instructions |

---

## Troubleshooting

| Issue | Solution |
|-------|----------|
| "Python not found" | Install Python and check "Add to PATH" |
| "Port in use" | Change 5180 to 5181 in start.bat |
| App doesn't open | Manually open http://localhost:5180 |

---

## System Requirements

- Python 3.10 or higher
- 500MB disk space
- Modern web browser (Chrome, Firefox, Edge)
