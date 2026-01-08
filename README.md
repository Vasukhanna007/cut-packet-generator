# Cut Packet Generator

A Streamlit web application for generating cut packet reports from Shopify order exports.

## Features

- Upload Shopify CSV exports
- Extract base products from product titles
- Normalize orders into components (Tops, Bottoms, Accessories)
- Filter by unfulfilled orders, date ranges, and order age
- Generate Excel reports with SUMIFS formulas for automatic totals
- Age-based filtering to show only orders older than X days
- Optional base product selection when using age filter

## Deployment

### Deploy to Streamlit Community Cloud (Recommended)

1. **Push to GitHub:**
   ```bash
   git remote add origin <your-github-repo-url>
   git branch -M main
   git push -u origin main
   ```

2. **Deploy on Streamlit Cloud:**
   - Go to [share.streamlit.io](https://share.streamlit.io)
   - Sign in with your GitHub account
   - Click "New app"
   - Select your repository and branch
   - Set Main file path to: `app.py`
   - Click "Deploy"

3. **Your app will be live at:** `https://your-app-name.streamlit.app`

### Alternative: Deploy to Other Platforms

- **Heroku:** Use the `Procfile` method
- **Railway:** Connect GitHub repo
- **Render:** Connect GitHub repo and set build command

## Local Development

```bash
# Create virtual environment
python3 -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate

# Install dependencies
pip install -r requirements.txt

# Run the app
streamlit run app.py
```

## Requirements

- Python 3.8+
- streamlit
- pandas
- openpyxl
- xlsxwriter
