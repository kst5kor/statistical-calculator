#!/bin/bash
cd "$(dirname "$0")"
source .venv/bin/activate
streamlit run "import streamlit as st.py" --server.port 5180 --server.headless true --browser.gatherUsageStats false
