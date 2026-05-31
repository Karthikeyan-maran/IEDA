# 📊 Intelligent Excel Data Analyst (IEDA)

An intelligent desktop application built with **Python** and **PyQt5** that transforms raw Excel data into actionable insights through interactive dashboards, ML-powered predictions, automated reporting, and LLM-based conversational analytics.

---

## ✨ Features

### 📁 Excel File Management
- Load and manage multiple Excel files (`.xlsx`, `.xls`) simultaneously
- Preview data (first 10 rows) or view the full sheet in a pop-up dialog
- Track loaded DataFrames with shape, column info, and metadata

### 📊 Interactive Dashboard
- **Chart Types**: Line, Bar, Pie, Scatter Plot, and Histogram
- **Data Filtering**: Filter data by column values before visualization
- **Multi-Column Selection**: Select multiple columns for comparison charts
- **Customizable**: Set chart titles, histogram bin counts, and more
- **Auto-Export**: Charts are automatically saved as PDF reports

### 🔧 Excel Operations
- **Merge DataFrames**: Combine two DataFrames using inner, left, right, or outer joins
- **Filter Data**: Filter by column values (supports `==`, `!=`, `>`, `<`, `>=`, `<=`, `contains`) or by row index
- **Search & Highlight**: Search across all columns for specific values
- All operation results displayed in pop-up dialogs and saved as downloadable files

### 📈 Data Analysis
- **Descriptive Statistics**: Mean, std, min/max, quartiles for numeric columns
- **Missing Values Analysis**: Total missing count and percentage per column
- **Value Counts**: Frequency distribution for categorical columns
- **PDF Report Generation**: Export analysis results as formatted PDF reports

### 🤖 ML Prediction
- **Linear Regression**: For numeric target prediction with R², MSE, MAE metrics
- **Decision Tree Classifier**: For categorical classification with accuracy and classification report
- **Feature Scaling**: Optional StandardScaler for feature normalization
- **Live Prediction**: Enter new values and get real-time predictions
- **LLM Explanation**: Ask the LLM to explain model results in plain language

### 💬 LLM Chat (Ollama Integration)
- Conversational AI powered by **Ollama** (using the `phi3` model)
- Optional Excel data context injection for data-aware responses
- Chat history with clear functionality

### 🤖 Automation Bot
- **Google Drive Integration**: Authenticate and fetch Excel files directly from Google Drive
- **Automated Operations**: Clear missing values, generate dashboard/analysis reports
- **Scheduling**: Schedule daily automation tasks at a specified time
- **Auto-Report Generation**: Generate and save PDF reports automatically

### 📂 Files & Reports Tab
- Central hub for all generated files (PDFs, Excel exports)
- **Download**: Save generated files to your local machine
- **Email**: Send reports directly via email (SMTP integration)

---

## 🛠️ Tech Stack

| Category | Technology |
|----------|-----------|
| **Language** | Python 3.9+ |
| **GUI Framework** | PyQt5 |
| **Data Processing** | Pandas, NumPy |
| **Visualization** | Matplotlib |
| **Machine Learning** | Scikit-learn (LinearRegression, DecisionTreeClassifier, StandardScaler) |
| **LLM Integration** | Ollama (phi3 model) |
| **Report Generation** | FPDF |
| **Cloud Storage** | PyDrive2 (Google Drive API) |
| **Email** | smtplib (SMTP) |
| **Scheduling** | schedule |
| **Theming** | Custom QSS Dark Theme |

---

## 📁 Project Structure

```
IEDA/
├── Excel Automation System/
│   ├── main_app.py            # Main application entry point & core UI
│   ├── AutomatedBotTab.py     # Google Drive automation & report generation
│   ├── analysis_tab.py        # Data analysis tab (descriptive stats, missing values)
│   ├── dashboard_tab.py       # Dashboard widget for chart generation
│   ├── prediction_tab.py      # ML prediction tab UI
│   ├── merge_tab.py           # DataFrame merge tab UI
│   ├── llm_chat_tab.py        # LLM chat tab UI
│   ├── llm_worker.py          # Background thread worker for Ollama LLM calls
│   ├── checked.png            # Custom checkbox icon (checked)
│   ├── unchecked.png          # Custom checkbox icon (unchecked)
│   ├── down_arrow.png         # Custom dropdown arrow icon
│   ├── .gitignore             # Git ignore rules
│   └── README.md              # This file
└── .gitignore                 # Root-level Git ignore rules
```

---

## 🚀 Getting Started

### Prerequisites
- **Python 3.9+** installed
- **Ollama** installed and running locally (for LLM features)
- **Google Cloud Project** with Drive API enabled (for Automation Bot)

### 1. Clone the Repository
```bash
git clone https://github.com/Karthikeyan-maran/IEDA.git
cd IEDA/Excel\ Automation\ System
```

### 2. Install Dependencies
```bash
pip install PyQt5 pandas numpy matplotlib scikit-learn ollama fpdf pydrive2 schedule
```

### 3. Set Up Ollama (for LLM Features)
```bash
# Install Ollama from https://ollama.com
# Pull the phi3 model
ollama pull phi3
```

### 4. Set Up Google Drive (for Automation Bot)
1. Create a project in [Google Cloud Console](https://console.cloud.google.com/)
2. Enable the **Google Drive API**
3. Download `client_secrets.json` and place it in the project root
4. See [PyDrive2 Quickstart](https://docs.iterative.ai/PyDrive2/quickstart/#authentication) for details

### 5. Run the Application
```bash
python main_app.py
```

---

## 🎨 UI Theme

The application features a **modern dark theme** built with custom QSS styling:
- Deep navy background (`#1a1a2e`)
- Purple accent buttons (`#3e2f5b`)
- Smooth border-radius on all components
- Custom checkbox and dropdown icons
- Dark-mode compatible Matplotlib charts

---

## 📧 Email Configuration

To use the email feature, update the sender credentials in `main_app.py` (`FilesTab.email_file` method):

```python
sender_email = "your_email@gmail.com"
sender_password = "your_app_password"
```

> **Note**: For Gmail, you must enable [App Passwords](https://support.google.com/accounts/answer/185833) in your Google Account security settings.

---

## 📋 Usage

1. **Load Data**: Click "Add Excel File(s)" to import your spreadsheets
2. **Explore**: Use the Dashboard tab to create charts and visualizations
3. **Analyze**: Switch to the Analysis tab for statistical summaries
4. **Predict**: Train ML models in the Prediction tab
5. **Chat**: Ask the LLM questions about your data in the LLM Chat tab
6. **Automate**: Set up scheduled tasks in the Automation Bot tab
7. **Export**: Download or email reports from the Files tab

---

## 🤝 Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

---

## 📝 License

This project is open source and available under the [MIT License](LICENSE).

---

## 👤 Author

**Karthikeyan Maran**  
- GitHub: [@Karthikeyan-maran](https://github.com/Karthikeyan-maran)
