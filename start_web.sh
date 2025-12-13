#!/bin/bash

# Start the RedLine Agent Web Interface

echo "Starting RedLine Agent Web Interface..."
echo ""
echo "Make sure you have:"
echo "  1. Installed dependencies: pip install -r requirements.txt"
echo "  2. Created a .env file with your API keys"
echo ""
echo "Opening browser to http://localhost:5000"
echo "Press Ctrl+C to stop the server"
echo ""

# Try to open browser (works on macOS and Linux with xdg-open)
sleep 2 && (open http://localhost:5000 2>/dev/null || xdg-open http://localhost:5000 2>/dev/null || echo "Please open http://localhost:5000 in your browser") &

python app.py








