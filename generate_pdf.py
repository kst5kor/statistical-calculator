"""Generate PDF setup guide for distribution."""

from fpdf import FPDF
from fpdf.enums import XPos, YPos

# Create PDF with proper margins
pdf = FPDF()
pdf.set_auto_page_break(auto=True, margin=15)
pdf.add_page()
pdf.set_margins(15, 15, 15)

# Title
pdf.set_font("Helvetica", "B", 20)
pdf.cell(
    0,
    15,
    "Statistical Process Capability Tool",
    new_x=XPos.LMARGIN,
    new_y=YPos.NEXT,
    align="C",
)
pdf.set_font("Helvetica", "", 12)
pdf.cell(
    0, 8, "Installation & Setup Guide", new_x=XPos.LMARGIN, new_y=YPos.NEXT, align="C"
)
pdf.ln(10)

# Features
pdf.set_font("Helvetica", "B", 14)
pdf.set_fill_color(230, 230, 230)
pdf.cell(0, 10, "Features", new_x=XPos.LMARGIN, new_y=YPos.NEXT, fill=True)
pdf.ln(3)
pdf.set_font("Helvetica", "", 11)
features = [
    "- Process Capability Indices (Cp, Cpk, Pp, Ppk)",
    "- Statistical hypothesis testing",
    "- Interactive visualizations with Plotly",
    "- Excel export with formatted reports",
    "- AI-powered Sigma Assistant",
]
for f in features:
    pdf.cell(0, 7, f, new_x=XPos.LMARGIN, new_y=YPos.NEXT)
pdf.ln(5)

# Prerequisites
pdf.set_font("Helvetica", "B", 14)
pdf.cell(0, 10, "Prerequisites", new_x=XPos.LMARGIN, new_y=YPos.NEXT, fill=True)
pdf.ln(3)
pdf.set_font("Helvetica", "", 11)
pdf.cell(
    0,
    7,
    "- Python 3.10 or higher (Download from python.org)",
    new_x=XPos.LMARGIN,
    new_y=YPos.NEXT,
)
pdf.cell(0, 7, "- pip (Python package manager)", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
pdf.ln(5)

# Installation Steps
pdf.set_font("Helvetica", "B", 14)
pdf.cell(0, 10, "Installation Steps", new_x=XPos.LMARGIN, new_y=YPos.NEXT, fill=True)
pdf.ln(3)

# Step 1
pdf.set_font("Helvetica", "B", 12)
pdf.cell(0, 8, "Step 1: Extract Files", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
pdf.set_font("Helvetica", "", 11)
pdf.cell(
    0, 6, "Extract the ZIP file to your computer.", new_x=XPos.LMARGIN, new_y=YPos.NEXT
)
pdf.ln(3)

# Step 2
pdf.set_font("Helvetica", "B", 12)
pdf.cell(
    0, 8, "Step 2: Open Terminal/Command Prompt", new_x=XPos.LMARGIN, new_y=YPos.NEXT
)
pdf.set_font("Helvetica", "", 11)
pdf.cell(0, 6, "Navigate to the extracted folder.", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
pdf.ln(3)

# Step 3
pdf.set_font("Helvetica", "B", 12)
pdf.cell(
    0, 8, "Step 3: Create Virtual Environment", new_x=XPos.LMARGIN, new_y=YPos.NEXT
)
pdf.set_font("Courier", "", 10)
pdf.set_fill_color(245, 245, 245)
pdf.cell(
    0,
    6,
    "Mac/Linux: python3 -m venv .venv && source .venv/bin/activate",
    new_x=XPos.LMARGIN,
    new_y=YPos.NEXT,
    fill=True,
)
pdf.cell(
    0,
    6,
    "Windows:   python -m venv .venv && .venv\\Scripts\\activate",
    new_x=XPos.LMARGIN,
    new_y=YPos.NEXT,
    fill=True,
)
pdf.ln(3)

# Step 4
pdf.set_font("Helvetica", "B", 12)
pdf.cell(0, 8, "Step 4: Install Dependencies", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
pdf.set_font("Courier", "", 10)
pdf.cell(
    0,
    6,
    "pip install -r requirements.txt",
    new_x=XPos.LMARGIN,
    new_y=YPos.NEXT,
    fill=True,
)
pdf.ln(3)

# Step 5
pdf.set_font("Helvetica", "B", 12)
pdf.cell(0, 8, "Step 5: Run the Application", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
pdf.set_font("Courier", "", 10)
pdf.cell(
    0,
    6,
    'streamlit run "import streamlit as st.py" --server.port 5180',
    new_x=XPos.LMARGIN,
    new_y=YPos.NEXT,
    fill=True,
)
pdf.ln(3)

# Step 6
pdf.set_font("Helvetica", "B", 12)
pdf.cell(0, 8, "Step 6: Open Browser", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
pdf.set_font("Helvetica", "", 11)
pdf.cell(
    0,
    6,
    "Open your browser and go to: http://localhost:5180",
    new_x=XPos.LMARGIN,
    new_y=YPos.NEXT,
)
pdf.ln(8)

# Quick Start
pdf.set_font("Helvetica", "B", 14)
pdf.cell(
    0,
    10,
    "Quick Start (After Initial Setup)",
    new_x=XPos.LMARGIN,
    new_y=YPos.NEXT,
    fill=True,
)
pdf.ln(3)
pdf.set_font("Helvetica", "", 11)
pdf.cell(
    0, 7, "- Mac/Linux: Double-click start.sh", new_x=XPos.LMARGIN, new_y=YPos.NEXT
)
pdf.cell(0, 7, "- Windows: Double-click start.bat", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
pdf.ln(5)

# Troubleshooting
pdf.set_font("Helvetica", "B", 14)
pdf.cell(0, 10, "Troubleshooting", new_x=XPos.LMARGIN, new_y=YPos.NEXT, fill=True)
pdf.ln(3)
pdf.set_font("Helvetica", "", 11)
pdf.cell(
    0,
    7,
    "Python not found: Install from python.org and add to PATH",
    new_x=XPos.LMARGIN,
    new_y=YPos.NEXT,
)
pdf.cell(
    0,
    7,
    "pip not found: Run python -m ensurepip --upgrade",
    new_x=XPos.LMARGIN,
    new_y=YPos.NEXT,
)
pdf.cell(
    0,
    7,
    "Port in use: Change to --server.port 5181",
    new_x=XPos.LMARGIN,
    new_y=YPos.NEXT,
)
pdf.cell(
    0,
    7,
    "Module not found: Re-run pip install -r requirements.txt",
    new_x=XPos.LMARGIN,
    new_y=YPos.NEXT,
)

# Save
pdf.output("Setup_Guide.pdf")
print("PDF created successfully: Setup_Guide.pdf")
