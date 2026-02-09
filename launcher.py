"""
Launcher script for PyInstaller packaging of the Statistical Calculator.
This starts the Streamlit server and opens the browser automatically.
"""

import os
import sys
import subprocess
import webbrowser
import time
import threading


def get_app_path():
    """Get the path to the main app file."""
    if getattr(sys, "frozen", False):
        # Running as compiled executable
        base_path = sys._MEIPASS
    else:
        # Running as script
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, "import streamlit as st.py")


def open_browser_delayed(url, delay=3):
    """Open browser after a delay to let server start."""
    time.sleep(delay)
    webbrowser.open(url)


def main():
    port = 5180
    url = f"http://localhost:{port}"

    # Start browser in background thread
    browser_thread = threading.Thread(target=open_browser_delayed, args=(url, 3))
    browser_thread.daemon = True
    browser_thread.start()

    # Get path to the streamlit app
    app_path = get_app_path()

    # Run streamlit
    from streamlit.web import cli as stcli

    sys.argv = [
        "streamlit",
        "run",
        app_path,
        "--server.port",
        str(port),
        "--server.headless",
        "true",
        "--browser.gatherUsageStats",
        "false",
    ]
    stcli.main()


if __name__ == "__main__":
    main()
