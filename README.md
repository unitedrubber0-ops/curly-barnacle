# EDI Tools Dashboard

A comprehensive Flask-based web application for managing EDI (Electronic Data Interchange) schedules with advanced features including conditional formatting, coverage analysis, and fluctuation tracking.

## Features

🎯 **Convert HTML → Weekly Excel**
- Transform HTML documents into structured Excel workbooks
- Automatic formatting and organization

📈 **Generate Coverage Report**
- Analyze schedules with three-color conditional formatting
- Green: Warehouse coverage
- Yellow: IT coverage  
- Red: Unmet demand
- Includes first date tracking for coverage changes

📊 **Compare EDI Fluctuations**
- Compare multiple EDI snapshots
- Identify trends and changes over time
- Bucket analysis by month

⚠️ **Critical Parts Report**
- Identify parts with unmet demand
- Export critical parts to Excel
- Configurable time windows

## Technology Stack

- **Backend**: Python 3.x, Flask
- **Data Processing**: pandas, openpyxl
- **Frontend**: HTML5, CSS3, JavaScript
- **Parsing**: BeautifulSoup4

## Project Structure

```
edi_web/
├── app.py                      # Main Flask application
├── edi.py                      # EDI processing utilities
├── requirements.txt            # Python dependencies
├── templates/                  # HTML templates
│   ├── layout.html            # Base template
│   ├── index.html             # Dashboard
│   ├── convert.html           # HTML converter
│   ├── coverage.html          # Coverage report
│   ├── fluctuations.html      # EDI comparison
│   └── critical_parts.html    # Critical parts
├── static/                     # Static assets
│   ├── css/style.css          # Bluish theme with floating buttons
│   └── images/logo.jpg        # Company logo
└── uploads/                    # User uploads
```

## Installation

1. **Clone the repository**:
```bash
git clone https://github.com/unitedrubber10-dot/fluffy-octo-giggle.git
cd edi_web
```

2. **Create virtual environment**:
```bash
python -m venv .venv
.\.venv\Scripts\Activate.ps1   # On Windows
source .venv/bin/activate       # On Linux/Mac
```

3. **Install dependencies**:
```bash
pip install -r requirements.txt
```

## Usage

1. **Start the Flask server**:
```bash
python app.py
```

2. **Open in browser**:
```
http://127.0.0.1:5000/
```

3. **Navigate to desired tool**:
   - Convert HTML files to Excel
   - Generate coverage reports with coloring
   - Compare EDI fluctuations
   - Export critical parts

## API Endpoints

| Method | Route | Purpose |
|--------|-------|---------|
| GET | `/` | Dashboard |
| GET/POST | `/convert` | HTML to Excel conversion |
| GET/POST | `/coverage` | Coverage report generation |
| GET/POST | `/fluctuations` | EDI fluctuation comparison |
| GET/POST | `/critical_parts` | Critical parts extraction |

## Configuration

### Coverage Report Parameters

- **sheet_name**: Excel sheet to analyze (default: "Schedule")
- **part_header**: Column header for part numbers (default: "PART")

### Critical Parts Parameters

- **window_weeks**: Number of weeks to analyze (default: 8)

## Color Coding

- 🟢 **Green**: Demand covered by warehouse stock
- 🟡 **Yellow**: Demand covered by IT (incoming transfers)
- 🔴 **Red**: Unmet demand

## Contributing

To contribute to this project:

1. Create a feature branch
2. Make your changes
3. Test thoroughly
4. Submit a pull request

## Support

For issues or questions, please contact the United Rubber Industries team.

## License

Proprietary - United Rubber Industries (I) PVT. LTD

---

**Company**: United Rubber Industries (I) PVT. LTD  
**Created**: December 2025  
**Version**: 1.0
