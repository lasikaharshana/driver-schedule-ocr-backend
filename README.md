# Driver Schedule OCR Backend

A Python Flask backend that uses Google Cloud Document AI to extract delivery schedules from board photos, automatically generate Excel truck load sheets in your preferred template, and provide results via API for integration with your mobile app.

---

## Features

- 📸 Upload a photo of a delivery schedule board
- 🤖 Extracts run numbers, drivers, trucks, and more using Document AI OCR
- 📊 Generates an Excel file per truck using your provided template
- 📱 Integrates seamlessly with the Android mobile app

---

## Requirements

- Python 3.8+
- Google Cloud account with Document AI enabled
- Google Service Account JSON key (for Document AI API)
- Your Excel truck load template (e.g., `Truck_Load_Record_Template.xlsx`)
- (Optional) Render.com or similar for production deployment

---

## Quick Start

```bash
# 1. Clone the repository
git clone https://github.com/lasikaharshana/driver-schedule-ocr-backend.git
cd driver-schedule-ocr-backend

# 2. Set up Python virtual environment and install dependencies
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt

# 3. Place your Google Cloud Document AI JSON key in the project directory,
#    or use Render’s secret file setting for production

# 4. Add your Excel template file (e.g., Truck_Load_Record_Template.xlsx) to the project root

# 5. Start the backend for local development

python app.py

# 6. To deploy, push to GitHub and redeploy on Render

