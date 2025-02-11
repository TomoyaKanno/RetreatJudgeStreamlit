from PyInstaller.utils.hooks import copy_metadata
import os
import streamlit

# Copy metadata as before.
datas = copy_metadata("streamlit")

# Locate the 'static' folder in the Streamlit package.
static_path = os.path.join(os.path.dirname(streamlit.__file__), "static")
# Append a tuple for the static folder: (source, destination relative path).
datas += [(static_path, "streamlit/static")]