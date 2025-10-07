# AMP Laydowns Automation

✅ **PRODUCTION READY** - Converts Excel marketing data into standardized AMP (Annual Marketing Plan) laydown presentations with pixel-perfect slides containing tables and charts.

## 🚀 Quick Start

### **Run the Production Script**
```bash
# Basic usage with your Excel data file
python scripts/excel_to_ppt_v1_071025.py --excel "your_data.xlsx" --template "template/Template_V4_FINAL_071025.pptx"

# Example with full paths
python scripts/excel_to_ppt_v1_071025.py \
    --excel "D:/Data/BulkPlanData.xlsx" \
    --template "D:/OneDrive - Publicis Groupe/work/(⭐) AMP Laydowns Automation/template/Template_V4_FINAL_071025.pptx"
```

### **Required Dependencies**
```bash
pip install pandas numpy python-pptx openpyxl
```

## 📁 Project Structure (Simplified October 7, 2025)

```
AMP Laydowns Automation/
├── .backup/                   # Project backups (1.5GB backup created Oct 7, 2025)
├── config/                    # Configuration files (6 files)
├── scripts/                   # Production script
│   └── excel_to_ppt_v1_071025.py  # Main production script (197KB, 3,928 lines)
├── template/                  # PowerPoint template
│   └── Template_V4_FINAL_071025.pptx  # Current production template
└── README.md                  # This file
```

## ✨ Key Features

### Core Capabilities
- ✅ **Excel to PowerPoint**: Automated presentation generation
- ✅ **Pixel-Perfect Formatting**: Professional template-based output (Template V4)
- ✅ **Table Splitting**: Automatic splitting for presentations with many campaigns (MAX 17 rows/slide)
- ✅ **Geography Processing**: Hierarchical geography data extraction (e.g., "Global | EMEA | MEA | Pakistan" → "Pakistan")
- ✅ **Multiple Media Types**: TV, Digital, OOH, Other with color coding
- ✅ **Chart Generation**: Funnel Stage, Media Type, Campaign Type pie charts
- ✅ **Enhanced TV Metrics**: GRP, Reach, Frequency calculations with sub-rows
- ✅ **Investment Sorting**: Slides ordered by total market investment
- ✅ **Professional Output**: Calibri fonts, precise positioning, branded colors

### Script Details (excel_to_ppt_v1_071025.py)
- **Lines of Code**: 3,928 lines
- **File Size**: 197 KB
- **Last Updated**: August 28, 2025
- **Features**:
  - Uses 'Plan - Geography' column for country extraction
  - MAX_ROWS_PER_SLIDE = 17 (optimized for template)
  - Pixel-perfect 2D coordinate system
  - Column indices corrected for production data (August 28 column fix)
  - Color-coded media types (TV: #71D48D, Digital: #FDF2B7, OOH: #FFBF00, Other: #B0D3FF)
  - Font sizes: Header 7.5pt, Body 7pt
  - Automatic campaign grouping and continuation slides

## 📊 Usage Examples

### Basic Usage
```bash
python scripts/excel_to_ppt_v1_071025.py \
    --excel "BulkPlanData.xlsx" \
    --template "template/Template_V4_FINAL_071025.pptx"
```

### With Full Paths
```bash
python scripts/excel_to_ppt_v1_071025.py \
    --excel "D:/Marketing/Data/2025/Q4/BulkPlanData.xlsx" \
    --template "D:/OneDrive - Publicis Groupe/work/(⭐) AMP Laydowns Automation/template/Template_V4_FINAL_071025.pptx"
```

### Command Line Arguments
- `--excel` : Path to Excel bulk plan data file (required)
- `--template` : Path to PowerPoint template file (required)

## 🔧 Technical Specifications

### Excel Column Mapping (Corrected August 28, 2025)
- **Column 10**: Plan - Geography (country extraction)
- **Column 17**: Plan - Brand
- **Column 20**: Media Type
- **Column 83**: **Campaign Name(s) (WAS 85)
- **Column 84**: **Campaign Type (WAS 86)
- **Column 95**: **Funnel Stage (WAS 97)
- **Column 71**: *Net Cost (WAS 73)
- **Column 55**: National GRP (WAS 56)
- **Column 104**: Reach 1+ (WAS 106)
- **Column 105**: Frequency (WAS 107)

### PowerPoint Slide Dimensions
- **Slide Size**: 10" × 5.625" (16:9 aspect ratio)
- **Table Position**: Left 0.184", Top 0.812", Width 9.299", Height 2.338"
- **Chart Position**: Y=3.300" (three pie charts side-by-side)
- **Title Bar**: Left 0.184", Top 0.308", Width 2.952", Height 0.370"

### Performance
- **Processing Speed**: ~5-10 minutes for 50 slides
- **Memory Usage**: ~500MB-1GB for typical datasets
- **Output Quality**: Production-ready, client-approved formatting

## 🔒 Important Notes

- **Backup Created**: 1.5GB full backup on October 7, 2025 (stored in `.backup/`)
- **Template**: Uses Template V4 (Template_V4_FINAL_071025.pptx)
- **Output**: Generated presentations saved to current directory or specified path
- **Logging**: Creates timestamped log files for troubleshooting
- **Column Fix**: August 28, 2025 update corrected 8 column indices for production data

## 📞 Support

For issues or questions:
1. Check the script logs (timestamped .log files)
2. Verify Excel column structure matches expected indices
3. Ensure template file exists and is accessible
4. Confirm all required Python packages are installed

---
**Version**: 1.0 (October 7, 2025)  
**Last Updated**: October 7, 2025  
**Status**: Production Ready ✅  
**Backup**: 1.5GB backup created October 7, 2025