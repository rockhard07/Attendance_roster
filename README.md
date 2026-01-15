# Attendance Analysis Tool - Streamlit Cloud Deployment

## 📊 Online Attendance Analysis Tool

A web-based application for converting PDF attendance reports into Excel files. Supports multiple departments and report types with intelligent data extraction.

### ✨ Features

- **📄 PDF to Excel Conversion**: Extract attendance data from PDF files and convert to Excel format
- **🏢 Multi-Department Support**: Works with Stations, OCC, and Train Operations departments
- **🚂 Train Operations Options**: 
  - **Trip Chart**: Extracts shift information and timings
  - **Roster**: Simple attendance extraction
- **👀 Data Preview**: View extracted data before downloading
- **⚡ Fast Processing**: Optimized for quick conversions
- **☁️ Cloud-Ready**: Hosted on Streamlit Cloud

### 🚀 Deployment on Streamlit Cloud

#### Prerequisites
- GitHub account with repository access
- Streamlit Cloud account (free tier available)

#### Step 1: Push to GitHub
```bash
cd git
git add .
git commit -m "Initial upload - Attendance Analysis Tool"
git push origin main
```

#### Step 2: Deploy on Streamlit Cloud
1. Visit [share.streamlit.io](https://share.streamlit.io)
2. Click "New app"
3. Select your GitHub repository
4. Choose branch: `main`
5. Set main file path: `git/app.py`
6. Click "Deploy"

The app will be live at: `https://your-username-app-name.streamlit.app`

### 📋 Supported PDF Formats

#### Stations & OCC
- Employee name, Personnel number, Scheduling row, Daily attendance codes
- Automatic Paid_Time detection from last column

#### Train Operations - Trip Chart
- Includes Shift and Shift Timings extraction
- Format: Column 4 contains shift info with timing (e.g., "SR-14\n05:00-13:00")
- Automatically splits into Shift and Shift_Timings columns

#### Train Operations - Roster
- Simple attendance extraction similar to Stations/OCC
- Employee roster with daily attendance records

### 🛠️ Local Development

```bash
# Install dependencies
pip install -r requirements.txt

# Run the app
streamlit run app.py
```

### 📁 Project Structure
```
git/
├── app.py                 # Main Streamlit application
├── pdf_extractor.py       # PDF extraction logic
│   ├── RosterExtractor    # Simple extraction (Stations/OCC/Train Ops Roster)
│   └── AttendancePDFExtractor # Trip Chart extraction (with Shift info)
├── requirements.txt       # Python dependencies
└── README.md             # This file
```

### 🔧 Configuration

No additional configuration needed for basic deployment. The app automatically:
- Detects Paid_Time columns
- Extracts shift information from PDF tables
- Handles multi-line cells and special characters

### 📊 Output Format

Excel files contain:
- **Employee**: Employee name
- **Personnel_Number**: Employee ID
- **Scheduling_Row**: Scheduling information
- **Shift**: (Trip Chart only) Shift name
- **Shift_Timings**: (Trip Chart only) Shift timing
- **Paid_Time**: Total paid hours (if detected)
- **Day_1 to Day_N**: Daily attendance codes

### 🐛 Troubleshooting

**Issue**: "No data found in PDF"
- Solution: Ensure PDF contains properly formatted tables with employee data

**Issue**: Shift information not extracted
- Solution: Verify 4th column contains shift data in Trip Chart PDFs

**Issue**: Preview shows incomplete data
- Solution: Check PDF file format and ensure all pages are readable

### 📝 Future Enhancements

- Batch PDF processing
- Multiple file upload
- Data analysis and reporting
- Employee performance metrics
- Night shift analysis
- Export to additional formats (CSV, JSON)

### 📄 License

MIT License - Feel free to use and modify

### 📧 Support

For issues and feature requests, please create an issue in the repository.

---

**Version**: 1.0  
**Last Updated**: January 2026  
**Status**: Beta
