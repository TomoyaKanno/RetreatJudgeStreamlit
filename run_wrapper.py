import sys
import os
import threading
import time
import webbrowser
from streamlit.web import cli as stcli

def launch_browser():
    # Wait a few seconds to give Streamlit time to start up.
    time.sleep(3)
    webbrowser.open("http://localhost:8501")

def get_resource_path(relative_path):
    """Get absolute path to resource, works for dev and for PyInstaller"""
    try:
        # PyInstaller creates a temp folder and sotres path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    
    return os.path.join(base_path, relative_path)

app_path = get_resource_path("app.py")

# Create a temporary directory for Streamlit's files if needed
if getattr(sys, 'frozen', False):
    os.environ['STREAMLIT_BROWSER_GATHER_USAGE_STATS'] = 'false'
    os.environ['STREAMLIT_SERVER_PORT'] = '8501'

# Set sys.argv to mimic running "streamlit run app.py" with additional options.
sys.argv = [
    "streamlit",
    "run",
    app_path,
    "--server.headless=true",
    "--global.developmentMode=false"
]

# Launch the browser in a separate thread.
threading.Thread(target=launch_browser).start()

# Start the Streamlit app.
sys.exit(stcli.main())
