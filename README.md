# 🔍 AI-Enhanced Excel Data Comparison Tool

An intelligent Python tool designed for the **HRMIS JoPayroll Backlog process** to compare consolidated payroll master lists with system-generated reports. This tool uses advanced rule-based matching combined with AI semantic similarity for borderline cases.

## 📋 What This Tool Does

This script compares two datasets containing employee information with the following columns:
- **Name**: Employee full name
- **Position**: Job title/position
- **TIN**: Tax Identification Number
- **Pay**: Net pay amount

The tool performs **10 different verification passes** and uses **AI semantic similarity** to make intelligent matching decisions for cases where traditional rules are insufficient.

## 📁 Files Included

- `Comparator.py`: Main Python script
- `requirements.txt`: All required Python packages
- `Simpler.xlsx`: Sample Excel file for testing
- `README.md`: This documentation

## 🛠️ Prerequisites

- **Python 3.8 or higher** (Python 3.10+ recommended)
- **Internet connection** (for first-time AI model download)
- **Excel file** with data in the correct format

## 📦 Installation Instructions

### Step 1: Check if Python is Installed

Open Command Prompt (cmd) and type:
```cmd
python --version
```

If Python is not installed, download it from [python.org](https://www.python.org/downloads/)

### Step 2: Navigate to the Project Folder

1. Open Command Prompt (Windows key + R, type `cmd`, press Enter)
2. Navigate to the folder containing the script files:
```cmd
cd "C:\Users\Lorenzo Bela\Downloads\Comparison Algorithm"
```
*Replace with your actual folder path*

### Step 3: Install Required Packages

**Option A: Install everything at once (Recommended)**
```cmd
pip install -r requirements.txt
```

**Option B: Install packages individually**
```cmd
pip install pandas>=1.5.0
pip install numpy>=1.21.0
pip install openpyxl>=3.0.0
pip install sentence-transformers>=2.2.2
pip install rapidfuzz>=3.0.0
pip install jellyfish>=1.0.0
pip install tqdm>=4.66.0
pip install joblib>=1.4.0
pip install textdistance>=4.6.2
pip install nameparser>=1.1.3
pip install python-stdnum>=1.19
pip install ftfy>=6.2.0
pip install unidecode>=1.3.7
pip install abydos>=0.5.0
pip install huggingface-hub>=0.24.0
pip install python-dotenv>=1.0.0
```

**Note**: Some packages are optional. If any package fails to install, the script will still work with reduced functionality.

## 📊 Excel File Format Requirements

Your Excel file should have data arranged as follows:

### Layout Example:
```
Column A | Column B | Column C | Column D | Column E | Column F | Column G | Column H | Column I
---------|----------|----------|----------|----------|----------|----------|----------|----------
Name1    | Pos1     | TIN1     | Pay1     |          | Name2    | Pos2     | TIN2     | Pay2
GARCIA,  | ADMIN    | 123-456- | 50000.00 |          | GARCIA,  | ADMIN    | 123456789| 50000.00
JUAN     | ASST III | 789      |          |          | JUAN C.  | ASST 3   |          |
...      | ...      | ...      | ...      |          | ...      | ...      | ...      | ...
```

### Default Column Ranges:
- **Dataset 1**: Columns A:D (Name, Position, TIN, Pay)
- **Dataset 2**: Columns F:I (Name, Position, TIN, Pay)
- **Column E**: Should be empty (separator)

### Important Notes:
- No headers are required (first row is treated as data)
- Both datasets should be on the same Excel sheet
- Empty rows are automatically ignored
- Names cannot be empty (will be excluded from comparison)

## 🚀 How to Run the Tool

### Step 1: Open Command Prompt
1. Press **Windows key + R**
2. Type `cmd` and press Enter
3. Navigate to your project folder:
```cmd
cd "C:\Users\Lorenzo Bela\Downloads\Comparison Algorithm"
```

### Step 2: Run the Script
```cmd
python Comparator.py
```

### Step 3: Follow the Interactive Prompts

1. **Accept Disclaimer**: Type `agree` to continue
2. **Enter File Path**: 
   - For the sample file: Type `.\Simpler.xlsx`
   - For your own file: Type the full path, e.g., `C:\Users\YourName\Desktop\MyData.xlsx`
3. **Choose Analysis Type**:
   - **Option 1**: AI-Enhanced comparison with default ranges (A:D and F:I) - **RECOMMENDED**
   - **Option 2**: AI-Enhanced comparison with custom column ranges
   - **Option 3-7**: Various testing and analysis options

### Example Run:
```cmd
C:\Users\Lorenzo Bela\Downloads\Comparison Algorithm> python Comparator.py

🔍 AI-Enhanced Excel Data Comparison Tool
============================================================
📋 Column Structure: Name, Position, TIN, Pay
📊 Default Ranges: A:D and F:I
🤖 AI Model: all-MiniLM-L6-v2 (for borderline cases)

===============================================================================
DISCLAIMER
===============================================================================
This script was developed by Lorenzo Bela under the supervision of the Data 
Warehouse and Analytics Division...
===============================================================================

Type 'agree' to continue: agree

Enter Excel file path (or press Enter for default): .\Simpler.xlsx

📋 Choose analysis type:
1. AI-Enhanced comparison with default ranges (A:D and F:I) - RECOMMENDED
2. AI-Enhanced comparison with manual column ranges
3. Test special character handling
4. Run multiple comparisons for accuracy testing
5. TIN-focused analysis (find all TIN mismatches)
6. Run comprehensive verification analysis
7. AI Model test (check if model loads correctly)

Enter your choice (1-7): 1
```

## 📊 Menu Options Explained

| Option | Description | When to Use |
|--------|-------------|-------------|
| **1** | AI-Enhanced comparison (default ranges) | **Most common use** - when your data is in A:D and F:I |
| **2** | AI-Enhanced comparison (custom ranges) | When your data is in different columns |
| **3** | Test special character handling | To verify the tool handles names with accents, special characters |
| **4** | Multiple comparison runs | For accuracy testing and validation |
| **5** | TIN-focused analysis | When you specifically need to find TIN mismatches |
| **6** | Comprehensive verification analysis | For detailed breakdown of all verification passes |
| **7** | AI model test | To verify AI components are working correctly |

## 📈 Understanding the Verification Process

The tool performs **10 verification passes**:

1. **Exact Name Match**: Perfect name matches
2. **First & Last Name Match**: Same first and last name components
3. **TIN Exact Match**: Identical TIN numbers
4. **TIN Format Match**: TIN matches after removing dashes/spaces
5. **Net Pay Exact Match**: Identical pay amounts (±0.01)
6. **Net Pay Close Match**: Pay within 5% difference
7. **Position Exact Match**: Identical position titles
8. **Position Normalized Match**: Similar positions after normalization
9. **Fuzzy Name Match**: High name similarity (≥80%)
10. **Comprehensive Score Match**: Overall similarity score ≥70%

### AI Decision Logic:
- **≥9 passes**: Automatic MATCH
- **≤5 passes**: Automatic MISMATCH
- **6-8 passes**: AI semantic analysis with intelligent thresholds
  - Name similarity ≥ 0.80
  - Position similarity ≥ 0.75

## 📄 Generated Reports

After running, the tool creates several reports:

| File Name | Description |
|-----------|-------------|
| `simpler_comparison_report.txt` | Summary report in text format |
| `simpler_comparison_report.xlsx` | Main Excel report with multiple sheets |
| `raw_data_comparison.xlsx` | Side-by-side comparison of raw data |
| `tin_mismatch_detailed_report.xlsx` | Detailed TIN mismatch analysis |
| `ai_validation_results_[timestamp].csv` | AI decision details for borderline cases |
| `missing_records.xlsx` | Records excluded during cleaning (if any) |
| `verification_analysis.xlsx` | Detailed verification pass results (Option 6) |

## 🎯 Accuracy Metrics

The tool provides comprehensive accuracy metrics:
- **Total Records Matched**: Number of successful matches
- **Perfect Matches**: Matches with identical pay and similar positions
- **Match Rate**: Percentage of records successfully matched
- **TIN Issues**: Number of TIN mismatches detected
- **AI Usage**: How many decisions required AI analysis

## 🔧 Troubleshooting

### Common Issues and Solutions:

#### "File not found" Error
```
❌ Error: File not found at [path]
```
**Solution**: 
- Check the file path is correct
- Use `.\filename.xlsx` for files in the same folder
- Use full path like `C:\Users\YourName\Desktop\file.xlsx`

#### Python Not Found
```
'python' is not recognized as an internal or external command
```
**Solution**:
- Try `py` instead of `python`
- Try `python3` instead of `python`
- Reinstall Python and check "Add to PATH" during installation

#### Package Installation Fails
```
ERROR: Could not install packages due to an EnvironmentError
```
**Solution**:
- Run Command Prompt as Administrator
- Use `pip install --user package_name`
- Update pip: `python -m pip install --upgrade pip`

#### AI Model Download Issues
```
❌ Failed to load AI model
```
**Solution**:
- Ensure internet connection is available
- The model will download automatically on first use (~90MB)
- Try running Option 7 to test AI model loading

#### Memory Issues with Large Files
- Close other applications
- Use a smaller sample of your data first
- Consider using Option 6 for verification analysis only

### Performance Tips:
- **Small files** (< 1000 records): All options work quickly
- **Medium files** (1000-10000 records): AI analysis may take a few minutes
- **Large files** (> 10000 records): Consider using verification analysis first

## 🔒 Important Security Notes

- This tool is designed for **HRMIS JoPayroll Backlog process**
- Handle confidential payroll data responsibly
- Do not share this script outside the Data Warehouse and Analytics Division
- Review the disclaimer before each use
- Ensure data files are stored securely

## 🆘 Getting Help

If you encounter issues:

1. **Check this README** for common solutions
2. **Try Option 7** to test if AI components are working
3. **Start with the sample file** (`Simpler.xlsx`) to verify the tool works
4. **Check your Excel file format** matches the requirements
5. **Contact the Data Warehouse and Analytics Division** for support

## 📝 Example Output

```
🎯 ACCURACY METRICS:
  • Total Records Matched: 245
  • Perfect Matches: 238
  • Match Rate: 89.1%
  • Perfect Match Rate: 97.1%
  • TIN Issues Found: 7 (See detailed report)

🤖 AI DECISION INSIGHTS:
  • Records processed with AI: 23
  • AI-confirmed matches: 18
  • AI-identified mismatches: 5
  • Borderline cases (6-8 rule passes): 12

✅ AI-Enhanced comparison completed successfully!
```

---

*This tool was developed by Lorenzo Bela under the supervision of the Data Warehouse and Analytics Division for the HRMIS JoPayroll Backlog process.*
