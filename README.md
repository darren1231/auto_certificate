# Auto Certificate Generator

A comprehensive GUI application for batch generating certificates using PowerPoint templates and Excel data sources. This tool allows users to create multiple certificate templates, map Excel columns to PowerPoint placeholders, and automatically generate personalized certificates.

## Features

- **Graphical User Interface**: User-friendly Tkinter-based interface for easy configuration
- **Multiple Templates**: Support for multiple certificate templates in a single session
- **Batch Processing**: Generate certificates for multiple recipients at once
- **Flexible Mapping**: Map Excel columns to PowerPoint text placeholders and table cells
- **Template Management**: Save and load template configurations as JSON files
- **Progress Logging**: Real-time logging of the generation process
- **Error Handling**: Robust error handling with detailed logging

## Prerequisites

- **Python 3.7+**
- **Microsoft PowerPoint** (required for PowerPoint automation)
- **Windows OS** (due to win32com dependency)

### Required Python Packages

- `pandas` - For Excel data processing
- `pywin32` - For PowerPoint automation
- `tkinter` - GUI framework (included with Python)

## Installation

1. **Clone or download the project files** to your local machine.

2. **Install required packages**:
   ```bash
   pip install pandas pywin32
   ```

3. **Ensure PowerPoint is installed** on your system.

## Usage

### Starting the Application

Run the main script:
```bash
python auto_certificate.py
```

### Interface Overview

The application has two main panels:

1. **Template Selection Panel (Left)**: Manage and select templates to execute
2. **Template Configuration Panel (Right)**: Configure selected template details

### Creating a Template

1. **Add a new template**:
   - Click the "+ 新增" (Add New) button
   - Enter a template name

2. **Configure template paths**:
   - **PPT Path**: Select your PowerPoint template file (.pptx)
   - **Excel Path**: Select your data source Excel file (.xlsx or .xls)

3. **Set up mapping rules**:
   - In the "Excel 欄位" (Excel Column) field, enter the column name from your Excel file
   - In the "PPT 標籤" (PPT Tag) field, enter the placeholder text in your PowerPoint template
   - Click "加入規則" (Add Rule) to add the mapping
   - Repeat for all required mappings

### Template Requirements

#### PowerPoint Template
- Use placeholder text like `{Name}`, `{Date}`, `{Award}` in your slides
- These placeholders will be replaced with actual data from Excel

#### Excel Data Source
- First row should contain column headers
- Each subsequent row represents one certificate recipient
- Column names must match exactly with the mapping rules (case-sensitive)

### Batch Processing

1. **Select templates to execute**:
   - Check/uncheck templates in the left panel
   - Use "全選/全不選" (Select All/None) to toggle all selections

2. **Start batch processing**:
   - Click "🚀 啟動勾選項目之批次任務" (Start Batch Process)
   - Monitor progress in the log area at the bottom

### Saving and Loading Configurations

- **Save Configuration**: Click "💾 儲存 Config" to save all templates as a JSON file
- **Load Configuration**: Click "📂 匯入 Config" to load previously saved templates

## Example Workflow

1. **Prepare your PowerPoint template**:
   - Create a slide with text like: "This certificate is awarded to {Name} for {Achievement} on {Date}"

2. **Prepare your Excel file**:
   ```
   Name        Achievement    Date
   John Doe    Excellence     2024-01-15
   Jane Smith  Leadership     2024-01-15
   ```

3. **Create mapping rules**:
   - Excel Column: "Name" → PPT Tag: "{Name}"
   - Excel Column: "Achievement" → PPT Tag: "{Achievement}"
   - Excel Column: "Date" → PPT Tag: "{Date}"

4. **Run the batch process** to generate personalized certificates for each row in Excel.

## Troubleshooting

### Common Issues

1. **"System clipboard occupied" error**:
   - Close applications that might be using the clipboard (LINE, clipboard managers, etc.)
   - Wait a moment and try again

2. **Excel column not found**:
   - Ensure column names in Excel exactly match the mapping rules
   - Check for extra spaces or special characters

3. **PowerPoint automation fails**:
   - Ensure Microsoft PowerPoint is properly installed
   - Try running as administrator

4. **No valid data found**:
   - Verify that Excel file has data rows
   - Check that mapping rules reference existing columns

### Debug Information

The application provides detailed logging in the bottom panel. Check the log for:
- Detected Excel columns
- Processing status for each certificate
- Error messages with specific details

## File Structure

```
project/
├── auto_certificate.py    # Main application script
├── README.md             # This documentation
├── Template/             # Directory for template files
│   ├── v2/
│   └── v3/
├── build/                # Build artifacts (if applicable)
└── *.json                # Configuration files
```

## Technical Details

- **Language**: Python 3
- **GUI Framework**: Tkinter
- **Automation**: win32com (COM automation for PowerPoint)
- **Data Processing**: pandas
- **Configuration Format**: JSON

## License

This project is provided as-is for educational and personal use.

## Contributing

Feel free to submit issues or pull requests for improvements.