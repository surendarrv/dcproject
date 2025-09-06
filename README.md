# Data Conversion Dashboard

An AI-driven Excel to Text Converter with a modern web interface and mainframe-style output preview.

## 🚀 Features

- **Modern Web UI**: Beautiful gradient interface with responsive design
- **Excel Processing**: Reads Excel files and detects MAPPING sheets automatically
- **Dynamic Column Detection**: Finds required columns (Begin, BETA Field Name, Mapping Instructions) dynamically
- **Mainframe Preview**: Authentic terminal-style output preview with green text on black background
- **File Upload/Download**: Easy file upload and text file download functionality
- **AI-Powered**: Claude-driven agentic control card generator

## 📋 Requirements

- Python 3.7+
- Flask
- pandas
- openpyxl

## 🛠️ Installation

1. Clone the repository:
```bash
git clone <your-repo-url>
cd python-file-project
```

2. Install dependencies:
```bash
pip install -r requirements.txt
```

3. Run the application:
```bash
python app.py
```

4. Open your browser and go to: `http://localhost:5000`

## 📁 Project Structure

```
python-file-project/
├── app.py                 # Flask web application
├── converter.py           # Excel to text conversion logic
├── templates/
│   ├── dashboard.html     # Modern web interface
│   └── index.html         # Alternative interface
├── Input1.xlsx           # Sample input file
├── input2.xlsx           # Sample input file
├── output1.txt           # Expected output format
├── output2.txt           # Expected output format
├── requirements.txt      # Python dependencies
└── README.md            # This file
```

## 🎯 How to Use

1. **Select File Type**: Choose "DEMO" (currently the only option)
2. **Upload Excel File**: Click "Select the raw file" and choose your .xlsx or .xls file
3. **Convert**: Click "Convert File" to process the Excel file
4. **Preview**: View the converted output in the mainframe-style preview window
5. **Download**: Click "Download TXT" to save the output as a text file

## 📊 Input Requirements

Your Excel file must contain:
- A sheet named "MAPPING" (case-insensitive)
- Columns: "Begin", "BETA Field Name", "Mapping Instructions for Programmer"

## 🔧 Technical Details

- **Column Detection**: Automatically finds required columns using flexible name matching
- **Field Name Generation**: Converts BETA field names to DEMO format (PER- → DEMO-, etc.)
- **Number Validation**: Handles float values in Begin column (1.0 → 0001)
- **Mapping Options**: Extracts mapping options from instructions (1. → 01, 2. → 02)

## 🎨 UI Features

- **Modern Design**: Purple-to-blue gradient background with clean white containers
- **Responsive Layout**: Works on desktop and mobile devices
- **Smooth Animations**: Hover effects and transitions
- **Mainframe Preview**: Authentic terminal look for output display
- **Real-time Statistics**: Shows total lines and field mappings count

## 📝 Sample Files

The project includes sample Excel files and their expected outputs:
- `Input1.xlsx` → `output1.txt`
- `input2.xlsx` → `output2.txt`

## 🤖 AI Integration

This application is powered by Claude AI and serves as an agentic control card generator for data transformation tasks.

## 📄 License

This project is part of a data conversion workflow system.

## 🆘 Support

For issues or questions, please check the code comments or create an issue in the repository.