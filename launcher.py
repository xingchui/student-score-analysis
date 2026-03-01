"""
Streamlit App Launcher for PyInstaller
This script allows packaging Streamlit apps as executable files.
"""
import sys
import os

# Get the directory where the executable is located
if getattr(sys, 'frozen', False):
    # Running as compiled executable
    base_path = os.path.dirname(sys.executable)
    app_path = os.path.join(base_path, 'student_analysis_app.py')
else:
    # Running as normal Python script
    base_path = os.path.dirname(os.path.abspath(__file__))
    app_path = os.path.join(base_path, 'student_analysis_app.py')

# Add base path to Python path
sys.path.insert(0, base_path)

# Change to the directory containing the app
os.chdir(base_path)

# Set environment variable for Streamlit
os.environ['STREAMLIT_CONFIG_DIR'] = os.path.join(base_path, '.streamlit')
os.environ['STREAMLIT_SERVER_FOLDER'] = base_path

# Launch Streamlit
if __name__ == '__main__':
    import subprocess
    subprocess.run([
        sys.executable, '-m', 'streamlit', 'run',
        app_path,
        '--server.port', '8501',
        '--server.headless', 'true',
        '--browser.gatherUsageStats', 'false'
    ])
