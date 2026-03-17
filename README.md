# Statistical Process Capability Tool

Single-characteristic Streamlit app for automotive dimensional capability analysis.

## What This App Does

- Enter part-by-part measurement data with `DMC / Serial Number` and `Value`
- Calculate capability metrics such as `Cp`, `Cpk`, `PPM`, required shift, and required tolerance
- Visualize process distribution, box plot, capability plots, and control chart
- Keep a local run history for repeated studies

This version is focused on automotive part manufacturing data and normal-process dimensional capability analysis.

## Project Structure

| File | Purpose |
|------|---------|
| `import streamlit as st.py` | Main Streamlit application |
| `requirements.txt` | Python dependencies |
| `start.sh` | Mac/Linux launcher |
| `start.bat` | Windows launcher |
| `Setup_Windows.bat` | Windows first-time setup |
| `Setup_Guide.pdf` | Printable user setup guide |
| `launcher.py` | Alternate launcher for packaged app flows |

## Local Run

### Mac/Linux

```bash
cd /path/to/py
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
streamlit run "import streamlit as st.py" --server.port 5180
```

Open: [http://localhost:5180](http://localhost:5180)

### Windows

First setup:

1. Install Python 3.10+ from [python.org](https://python.org)
2. Make sure `Add Python to PATH` is checked
3. Run `Setup_Windows.bat`
4. Open [http://localhost:5180](http://localhost:5180)

After setup:

- Run `start.bat`

## Team Sharing

Best practice for a team is to keep this `py/` folder in a shared Git repository and use one of these models:

1. Internal shared Streamlit deployment
   Best for quality/manufacturing teams that just need a browser link.
2. Git + local run
   Best for engineering teams that want to review formulas and make controlled updates.
3. Packaged desktop build
   Best for shop-floor users who should not install Python.

Recommended default:

- Keep one reviewed `main` branch
- Use pull requests for formula or UI changes
- Treat capability formulas as controlled logic
- Test with a known sample dataset before release

## Suggested Team Workflow

1. Clone the repo
2. Create and activate `.venv`
3. Install `requirements.txt`
4. Run the app locally
5. Validate against a known sample part dataset
6. Merge only reviewed changes into the shared branch

## Deployment Notes

For an internal shared deployment, run:

```bash
streamlit run "import streamlit as st.py" --server.port 5180 --server.address 0.0.0.0
```

Then publish the host URL inside your company network.

## Troubleshooting

| Issue | Solution |
|-------|----------|
| Python not found | Reinstall Python and add it to PATH |
| Port 5180 already in use | Change the port in the launch command or start script |
| Browser does not open | Open [http://localhost:5180](http://localhost:5180) manually |
| App seems stale after edits | Stop and restart Streamlit |

## Requirements

- Python 3.10 or higher
- Modern browser https://statistical-calculator-sw83whcgzqniq4ndfpexgt.streamlit.app/
- Local network access only if you plan to share it on an internal server
