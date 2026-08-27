# Heart Disease Prediction & Recommendation System

A Machine Learning and Streamlit application for Heart Disease Risk Prediction and Interactive Analytics.

## Features
- **Heart Disease Prediction**: Uses a trained K-Nearest Neighbors (KNN) model with feature scaling to assess the risk of heart disease based on patient clinical parameters.
- **Interactive Web Interface**: Streamlit UI for user-friendly medical data input and immediate prediction feedback.
- **FastAPI & Streamlit Services**: FastAPI backend and Streamlit interface utilizing TF-IDF and TMDB API.

## Project Structure
- `heart1.py` - Streamlit application for Heart Disease Prediction.
- `KNN_heart.pkl` - Trained K-Nearest Neighbors model.
- `scaler.pkl` - Pre-fitted StandardScaler for feature normalization.
- `columns.pkl` - Expected feature column structure for inference.
- `main.py` - FastAPI backend application.
- `app.py` - Streamlit frontend interface.
- `requirements.txt` - Python project dependencies.

## Installation & Setup

1. **Clone the repository:**
   ```bash
   git clone https://github.com/monusharma674/Heart-Diseases-Predication.git
   cd Heart-Diseases-Predication
   ```

2. **Create and activate a virtual environment:**
   ```bash
   python -m venv venv
   # On Windows:
   venv\Scripts\activate
   # On macOS/Linux:
   source venv/bin/activate
   ```

3. **Install dependencies:**
   ```bash
   pip install -r requirements.txt
   ```

4. **Environment Variables (Optional):**
   Copy `.env.example` to `.env` and set your API keys if needed:
   ```bash
   cp .env.example .env
   ```

## Running the Application

### Heart Disease Prediction App:
```bash
streamlit run heart1.py
```

### FastAPI Backend & Streamlit App:
```bash
# Run API server
uvicorn main:app --reload

# Run Streamlit Frontend
streamlit run app.py
```
