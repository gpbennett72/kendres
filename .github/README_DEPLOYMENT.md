# Deployment Guide

This app can be deployed to various platforms. Here are the recommended options:

## Railway (Recommended)

1. Sign up at [railway.app](https://railway.app)
2. Create a new project
3. Connect your GitHub repository
4. Add environment variables:
   - `OPENAI_API_KEY` or `ANTHROPIC_API_KEY`
   - `GOOGLE_CREDENTIALS_PATH` (if using Google Docs)
   - `FLASK_SECRET_KEY` (optional, will be auto-generated)
5. Railway will auto-detect Python and deploy

## Render

1. Sign up at [render.com](https://render.com)
2. Create a new Web Service
3. Connect your GitHub repository
4. Settings:
   - Build Command: `pip install -r requirements.txt`
   - Start Command: `gunicorn app:app` (or use the provided start script)
5. Add environment variables in the Render dashboard

## Fly.io

1. Install flyctl: `curl -L https://fly.io/install.sh | sh`
2. Run `fly launch` in the project directory
3. Follow the prompts to configure
4. Set secrets: `fly secrets set OPENAI_API_KEY=your_key`

## Environment Variables

Create a `.env` file (or set in your platform's dashboard):

```env
OPENAI_API_KEY=your_key_here
# OR
ANTHROPIC_API_KEY=your_key_here

DEFAULT_AI_PROVIDER=openai
DEFAULT_MODEL=gpt-4

GOOGLE_CREDENTIALS_PATH=credentials.json
FLASK_SECRET_KEY=your_secret_key_here
```

## Notes

- The app requires Python 3.11 or 3.12
- All dependencies are listed in `requirements.txt`
- For Google Docs support, you'll need OAuth2 credentials from Google Cloud Console

