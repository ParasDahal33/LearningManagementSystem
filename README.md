# Canvas Quiz Uploader

A Streamlit tool to parse DOCX quizzes and upload them to Canvas LMS.

## Setup

1. **Install dependencies:**
   ```bash
   pip install -r requirements.txt
   ```

2. **Configuration (Optional):**
   You can store secrets in `.streamlit/secrets.toml` (in the project root) or use the UI login.
   
   Example `.streamlit/secrets.toml`:
   ```toml
   CANVAS_BASE_URL = "https://<your-domain>.instructure.com/api/v1"
   CANVAS_TOKEN = "your_canvas_token"
   OPENAI_API_KEY = "sk-..."
   ```

## Running the App

This application is built with Streamlit. You must run it using the `streamlit` CLI, not `python`.

```bash
streamlit run app.py
```

The app will open in your default web browser.