# Quick Start Guide - RedLine Web Interface

## Step 1: Install Dependencies

```bash
pip install -r requirements.txt
```

## Step 2: Set Up API Keys

Create a `.env` file in the project root:

```env
OPENAI_API_KEY=your_openai_api_key_here
# OR
ANTHROPIC_API_KEY=your_anthropic_api_key_here

DEFAULT_AI_PROVIDER=openai
DEFAULT_MODEL=gpt-4
```

## Step 3: Start the Web Server

### Option A: Using the startup script
```bash
./start_web.sh
```

### Option B: Direct Python command
```bash
python app.py
```

## Step 4: Open Your Browser

Navigate to: **http://localhost:5000**

## Using the Web Interface

### For Word Documents:

1. **Upload Files Tab**:
   - Click "Word Document" tab (default)
   - Upload your Word document (.docx) with tracked changes
   - Upload your legal playbook file (.txt)

2. **Configure AI**:
   - Select AI provider (OpenAI or Anthropic)
   - Choose model (GPT-4, Claude, etc.)
   - Choose whether to create a summary document

3. **Process**:
   - Click "Analyze Only" to see analysis without modifying the document
   - Click "Process & Add Comments" to add comments to the document

4. **View Results**:
   - See summary of redlines found
   - Review detailed analysis for each redline
   - Download the processed document with comments
   - Download summary document (if created)

### For Google Docs:

1. **Google Doc Tab**:
   - Click "Google Doc" tab
   - Enter Google Doc ID or full URL
   - Paste your legal playbook content
   - Configure AI settings
   - Click "Process Google Doc"

2. **Note**: First-time Google Docs usage requires OAuth2 setup:
   - Download credentials from Google Cloud Console
   - Place in project directory as `credentials.json`
   - The app will guide you through OAuth on first use

## Sample Playbook

A sample playbook (`sample_playbook.txt`) is included. You can use it as a template or modify it with your own legal guidelines.

## Troubleshooting

### "API Key not found" error
- Make sure your `.env` file exists and contains the correct API key
- Check that the key name matches exactly (OPENAI_API_KEY or ANTHROPIC_API_KEY)

### "Files not found" error
- Make sure you've uploaded both the document and playbook files
- Try refreshing the page and uploading again

### Google Docs authentication issues
- Make sure `credentials.json` is in the project directory
- Check that Google Docs API is enabled in your Google Cloud project
- Delete `token.pickle` and re-authenticate if needed

### Port already in use
- Change the port in `app.py` (last line): `app.run(..., port=5001)`
- Or stop the process using port 5000

## Tips

- Start with "Analyze Only" to test without modifying documents
- Use the sample playbook to understand the format
- Check the console output for detailed error messages
- Large documents may take several minutes to process








