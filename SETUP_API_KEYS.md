# Setting Up API Keys for RedLine Agent

The RedLine Agent requires API keys from either OpenAI or Anthropic to analyze redlines using AI.

## Option 1: OpenAI (Recommended for Testing)

1. **Get an API Key:**
   - Go to https://platform.openai.com/api-keys
   - Sign up or log in to your OpenAI account
   - Click "Create new secret key"
   - Copy the key (you'll only see it once!)

2. **Add to .env file:**
   Create a `.env` file in the project root with:
   ```env
   OPENAI_API_KEY=sk-your-actual-api-key-here
   DEFAULT_AI_PROVIDER=openai
   DEFAULT_MODEL=gpt-4
   ```

## Option 2: Anthropic (Claude)

1. **Get an API Key:**
   - Go to https://console.anthropic.com/
   - Sign up or log in
   - Navigate to API Keys section
   - Click "Create Key"
   - Copy the key

2. **Add to .env file:**
   Create a `.env` file in the project root with:
   ```env
   ANTHROPIC_API_KEY=sk-ant-your-actual-api-key-here
   DEFAULT_AI_PROVIDER=anthropic
   DEFAULT_MODEL=claude-3-opus-20240229
   ```

## Quick Setup

Run this command to create a `.env` file template:

```bash
cat > .env << 'EOF'
# Choose one provider (OpenAI or Anthropic)
OPENAI_API_KEY=your_openai_key_here
# OR
# ANTHROPIC_API_KEY=your_anthropic_key_here

# Default settings
DEFAULT_AI_PROVIDER=openai
DEFAULT_MODEL=gpt-4

# Google Docs (optional, only if using Google Docs)
# GOOGLE_CREDENTIALS_PATH=credentials.json
EOF
```

Then edit the `.env` file and add your actual API key.

## Cost Considerations

- **OpenAI GPT-4**: ~$0.03 per 1K input tokens, ~$0.06 per 1K output tokens
- **OpenAI GPT-3.5 Turbo**: Much cheaper, ~$0.0015 per 1K tokens
- **Anthropic Claude**: Varies by model, generally competitive with GPT-4

For testing, GPT-3.5 Turbo is the most cost-effective option.

## Verify Your Setup

After creating your `.env` file, you can verify it's working by:

1. Starting the web server: `python app.py`
2. Uploading a test document
3. The server will show an error if the API key is invalid

## Security Note

⚠️ **Never commit your `.env` file to git!** It's already in `.gitignore` for your protection.








