# Software Version Analysis Tool

A Python-based tool that analyzes installed software versions on Windows systems and compares them against the latest available versions from multiple package managers and web sources.

## 🚀 Features

- **Multi-Source Version Checking**: Queries multiple sources to find the latest software versions:
  - Chocolatey package manager
  - Windows Package Manager (winget)
  - Web scraping from official software repositories
- **Excel Integration**: Reads software data from Excel files and exports results to Excel
- **Intelligent Version Comparison**: Uses semantic versioning to accurately compare software versions
- **Comprehensive Reporting**: Generates detailed reports with status indicators (Updated/Outdated/Unknown)
- **Error Handling**: Robust error handling for network issues and missing software

## 📋 Prerequisites

### Required Software

- **Python 3.7+**
- **Chocolatey** (for package manager queries)
- **Windows Package Manager (winget)** (for Microsoft Store queries)

### Python Dependencies

Install the required Python packages:

```bash
pip install pandas openpyxl requests packaging
```

## 📁 File Structure

```
project/
├── delv3.py                          # Main analysis script
├── 2024_06_20_Software_Analysis_All.xlsx  # Input Excel file
├── ACC_IMS_Third_Report.xlsx         # Output report (generated)
└── README.md                         # This file
```

## 🔧 Installation

1. **Clone or download** this repository to your local machine
2. **Install Python dependencies**:
   ```bash
   pip install pandas openpyxl requests packaging
   ```
3. **Ensure Chocolatey is installed** (if not already):
   ```powershell
   Set-ExecutionPolicy Bypass -Scope Process -Force; [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072; iex ((New-Object System.Net.WebClient).DownloadString('https://community.chocolatey.org/install.ps1'))
   ```
4. **Verify winget is available** (comes with Windows 10/11 by default)

## 📊 Input Format

The tool expects an Excel file with the following columns:

- **DisplayName**: Software name
- **DisplayVersion**: Currently installed version

### Example Input Data:

| DisplayName        | DisplayVersion |
| ------------------ | -------------- |
| Google Chrome      | 120.0.6099.109 |
| Mozilla Firefox    | 121.0          |
| Visual Studio Code | 1.85.1         |

## 🚀 Usage

1. **Prepare your input file**: Ensure your Excel file is named `2024_06_20_Software_Analysis_All.xlsx` and contains the required columns
2. **Run the script**:
   ```bash
   python delv3.py
   ```
3. **Check the output**: The tool will generate `ACC_IMS_Third_Report.xlsx` with the analysis results

## 📈 Output Format

The tool generates an Excel report with the following columns:

| Column            | Description                              |
| ----------------- | ---------------------------------------- |
| Software Name     | Name of the software                     |
| Installed Version | Currently installed version              |
| Latest Version    | Latest available version found           |
| Status            | Update status (Updated/Outdated/Unknown) |

### Status Meanings:

- **Updated**: Installed version is current or newer
- **Outdated**: A newer version is available
- **Unknown**: Unable to determine version status

## 🔍 How It Works

### 1. Data Loading

- Reads software data from Excel file
- Filters out invalid entries (undefined, empty values)
- Removes duplicates based on software name

### 2. Version Source Checking (in order of preference)

1. **Chocolatey**: Queries the Chocolatey package repository
2. **Winget**: Searches Windows Package Manager
3. **Web Scraping**: Checks official software websites

### 3. Version Comparison

- Uses semantic versioning for accurate comparisons
- Handles different version formats (x.x, x.x.x, etc.)
- Provides fallback mechanisms for version parsing

### 4. Report Generation

- Creates comprehensive Excel report
- Provides summary statistics
- Includes all software with their status

## ⚙️ Configuration

### Customizing Input File

To use a different input file, modify line 8 in `delv3.py`:

```python
file_path = "your_file_name.xlsx"
```

### Adjusting Timeouts

Modify timeout values in the script for slower networks:

```python
timeout=10  # Increase for slower connections
```

### Changing Output File

Modify line 218 in `delv3.py`:

```python
output_path = "your_output_file.xlsx"
```

## 🛠️ Troubleshooting

### Common Issues

1. **"choco not found"**

   - Install Chocolatey: https://chocolatey.org/install
   - Ensure it's in your system PATH

2. **"winget not found"**

   - Update Windows to get winget
   - Or install from Microsoft Store

3. **Network timeouts**

   - Increase timeout values in the script
   - Check your internet connection

4. **Excel file not found**
   - Ensure the input file is in the same directory as the script
   - Check the filename matches exactly

### Performance Tips

- The script includes a 0.5-second delay between checks to avoid overwhelming servers
- For large software lists, consider running during off-peak hours
- Some web sources may be rate-limited

## 📊 Sample Output

```
[1/150] Checking Google Chrome (installed: 120.0.6099.109)...
    Checking multiple sources for Google Chrome...
      Chocolatey: 120.0.6099.130
[2/150] Checking Mozilla Firefox (installed: 121.0)...
    Checking multiple sources for Mozilla Firefox...
      Winget: 121.0.1

📊 Summary:
  Updated: 45
  Outdated: 23
  Unknown: 82

✅ Report saved to ACC_IMS_Third_Report.xlsx
```

## 🤝 Contributing

Feel free to submit issues and enhancement requests!

## 📄 License

This project is open source and available under the [MIT License](LICENSE).

## 📞 Support

For issues or questions:

1. Check the troubleshooting section above
2. Review the error messages in the console output
3. Ensure all prerequisites are properly installed

---

**Note**: This tool is designed for Windows systems and requires administrative privileges for some package manager operations.
