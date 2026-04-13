#!/usr/bin/env python3
"""
Launch script for the Streamlit CostHead Report Generator
"""

import subprocess
import sys
import os
from pathlib import Path

def main():
    """Launch the Streamlit app"""

    # Get the directory where this script is located
    script_dir = Path(__file__).parent
    app_file = script_dir / "streamlit_app.py"

    # Check if the app file exists
    if not app_file.exists():
        print(f"❌ Error: {app_file} not found!")
        sys.exit(1)

    print("🚀 Starting CostHead Report Generator...")
    print(f"📁 App directory: {script_dir}")
    print(f"📄 App file: {app_file}")
    print("\n" + "="*50)
    print("🌐 The app will open in your default web browser")
    print("📊 Access the app at: http://localhost:8501")
    print("⏹️  Press Ctrl+C to stop the server")
    print("="*50 + "\n")

    try:
        # Use the same Python as this script (works with venv, conda, etc. on Windows and Unix)
        result = subprocess.run(
            [
                sys.executable,
                "-m",
                "streamlit",
                "run",
                str(app_file),
                "--server.port",
                "8501",
                "--server.address",
                "localhost",
                "--browser.gatherUsageStats",
                "false",
            ],
            cwd=script_dir,
        )
        if result.returncode:
            sys.exit(result.returncode)
    except KeyboardInterrupt:
        print("\n👋 Shutting down CostHead Report Generator...")
    except Exception as e:
        print(f"❌ Error launching Streamlit: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()
