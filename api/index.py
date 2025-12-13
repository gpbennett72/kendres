import sys
import os
import traceback

# Setup paths
parent_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, parent_dir)
os.chdir(parent_dir)
os.environ['VERCEL'] = '1'

# Import Flask app with error handling
# First, try to ensure lxml is available (python-docx requires it)
try:
    from lxml import etree
    print("✓ lxml.etree imported successfully", flush=True)
except ImportError as lxml_error:
    print(f"⚠ WARNING: lxml.etree import failed: {lxml_error}", flush=True)
    print("⚠ This may cause issues with python-docx. Attempting to continue...", flush=True)

# Check pydantic-core availability - critical for OpenAI/Anthropic SDKs
# This is a known issue with pydantic-core binary wheels in Vercel Python 3.12
try:
    import pydantic_core._pydantic_core
    print("✓ pydantic_core._pydantic_core imported successfully", flush=True)
except (ImportError, AttributeError) as pydantic_error:
    print(f"❌ CRITICAL: pydantic_core._pydantic_core import failed: {pydantic_error}", flush=True)
    print("❌ This will cause failures when using OpenAI/Anthropic SDKs", flush=True)
    # Try to diagnose the issue
    try:
        import pydantic_core
        print(f"⚠ pydantic_core module found at: {pydantic_core.__file__}", flush=True)
        import os
        core_dir = os.path.dirname(pydantic_core.__file__)
        print(f"⚠ pydantic_core directory: {core_dir}", flush=True)
        # Check for extension files
        try:
            files = [f for f in os.listdir(core_dir) if '_pydantic_core' in f]
            print(f"⚠ Files with '_pydantic_core' in name: {files}", flush=True)
        except Exception as e:
            print(f"⚠ Could not list directory: {e}", flush=True)
    except ImportError as e2:
        print(f"❌ pydantic_core module not found: {e2}", flush=True)

try:
    from app import app
except Exception as e:
    # If import fails, create a minimal Flask app that shows the error
    from flask import Flask, jsonify
    error_app = Flask(__name__)
    
    error_msg = str(e)
    error_tb = traceback.format_exc()
    
    @error_app.route('/', defaults={'path': ''})
    @error_app.route('/<path:path>')
    def error_handler(path):
        # Return HTML error page for better readability
        html = f"""
        <!DOCTYPE html>
        <html>
        <head>
            <title>Application Error</title>
            <style>
                body {{ font-family: monospace; padding: 20px; background: #f5f5f5; }}
                .error-box {{ background: white; padding: 20px; border-radius: 5px; border-left: 4px solid #e00; }}
                h1 {{ color: #e00; margin-top: 0; }}
                pre {{ background: #f5f5f5; padding: 15px; border-radius: 3px; overflow-x: auto; }}
            </style>
        </head>
        <body>
            <div class="error-box">
                <h1>Application Failed to Load</h1>
                <h2>Error Message:</h2>
                <pre>{error_msg}</pre>
                <h2>Traceback:</h2>
                <pre>{error_tb}</pre>
            </div>
        </body>
        </html>
        """
        return html, 500
    
    app = error_app
