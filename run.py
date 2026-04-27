"""
CIS School System - Entry Point
=================================
Run this file to start the server.

Usage (from the project root):
    python run.py
"""

from app import create_app

app = create_app()

if __name__ == "__main__":
    print("\n" + "=" * 55)
    print("  CIS School Management System")
    print("  Running at: http://127.0.0.1:5000")
    print("  Press CTRL+C to stop")
    print("=" * 55 + "\n")
    app.run(debug=True, host="0.0.0.0", port=5000)
