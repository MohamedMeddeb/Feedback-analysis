# Feedback Analysis AI

📊 Your intelligent feedback insights assistant

## ✨ Features

🔍 **Sentiment Analysis** - Understand emotional tone in feedback  
📈 **Category Scoring** - Automatic skill categorization (Business, Technical, Leadership)  
🌍 **Multilingual Support** - Works with English and French feedback  
📊 **Visual Analytics** - Interactive charts and improvement tracking  

## ⚠️ Important Requirements

🐍 **Python 3.8-3.12 ONLY**  
(spaCy installation currently requires Python ≤ 3.12)

## 📋 How to Use

### Main Analysis
1. Upload Excel/CSV feedback files
2. Enter the employee's Feedback ID
3. View comprehensive sentiment and category analysis

### Results Breakdown
📊 **Scorecard** - Overall performance metrics  
📈 **Trend Charts** - Skill improvement over time  
💬 **Key Comments** - Most significant feedback excerpts  

## 🛠️ Technical Highlights

Feature | Description
---|---
NLP Engine | Uses spaCy and TextBlob for advanced text analysis
Smart Categorization | Hybrid model combining keywords and semantic similarity
Translation | Automatic French-to-English translation of feedback
Visualization | Interactive Matplotlib charts in Tkinter GUI

## 🚀 Getting Started

```bash
# Clone the repository
git clone https://github.com/MohamedMeddeb/Feedback-analysis.git

# Install dependencies (Python 3.8-3.12 required)
pip install pandas textblob googletrans langdetect spacy matplotlib seaborn

# Download language model
python -m spacy download en_core_web_md

# Run the application
python feedback_analyzer.py
