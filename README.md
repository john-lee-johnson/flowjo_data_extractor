# Flow Cytometry Data Processor

A graphical user interface (GUI) tool for processing and analyzing flow cytometry data. This program simplifies the process of combining and analyzing data tables from FlowJo with sample and group information in a plate map format.

## Features

- Load and process FlowJo exported data
- Organize samples and group into plate maps
- Filter output by samples or groups
- Multiple analysis options:
  - Individual replicates
  - Mean with Standard Deviation (SD)
  - Mean with Standard Error of Mean (SEM)
- Flexible export options:
  - Standard (for grouped Prism format) or single-row format (for XY tables)
  - Copy to clipboard
  - Save to CSV
  - Optional headers

### Required File Formats

#### Sample and Group Maps
- Excel files (.xlsx or .xls)
- Format should be a plate layout with:
  - Empty cell in top-left corner
  - Column numbers (1-12) across the top
  - Row letters (A-H) down the left side
  - Sample/Group names in the corresponding wells

Example map layout:
```
   1    2    3    4    ...
A  WT   WT   KO   KO   
B  WT   WT   KO   KO   
...
```

#### FlowJo Data
- Excel file (.xlsx or .xls) exported from FlowJo
- Sample names must include well positions in the format: `*_[WELL].fcs`
  - Example: `Sample_A01.fcs`
- First column should contain sample names
  - Sample names don't need to match the sample map, but the wells need to match
- Additional columns contain measurements

## Installation
1. Make a folder to hold the program (e.g., Documents/scripts/)
2. Open terminal (using Launchpad or Applications folder)
3. Navigate to the folder you just created
   ```bash
   cd ~/Documents/scripts/
   ```
4. Clone the repository:
   ```bash
   git clone john-lee-johnson/flowjo-data-extractor
   ```
5. Install the required packages:
   ```bash
   pip3 install PyQt6 pandas numpy scipy pyperclip openpyxl xlrd
   ```

## Usage

1. With the terminal open, go to the folder you downloaded
   ```bash
   cd ~/Documents/scripts/flowjo-data-extractor
   ```
2. Start the program:
   ```bash
   python3 flowjo_data_extraction.py
   ```

3. Load your data files:
   - **Sample Map**: Excel file mapping well positions to sample names
   - **Group Map**: Excel file mapping well positions to group names
   - **FlowJo Data**: Excel file exported from FlowJo

### Analysis Options

1. **Individual Replicates**
   - Shows all individual data points
   - Groups replicate measurements together

2. **Mean & SD**
   - Calculates mean and standard deviation for each group
   - Displays results in alternating Mean/SD columns

3. **Mean & SEM**
   - Calculates mean and standard error of the mean for each group
   - Displays results in alternating Mean/SEM columns

### Export Options

1. **Standard Format**
   - Samples in rows
   - Groups in columns
   - Good for "grouped" table in Prism
   - Optional headers
   - Example:
     ```
     Sample  Group1  Group2  Group3
     A       10.5    15.2    12.8
     B       11.2    14.8    13.1
     ```

2. **Single Row Format**
   - All data in a single row
   - Good for "XY" table in Prism
   - Sample_Group combinations as column headers
   - Example:
     ```
     A_Group1  A_Group2  B_Group1  B_Group2
     10.5      15.2      11.2      14.8
     ```

## Troubleshooting

Common issues and solutions:

1. **File Loading Errors**
   - Ensure files are in the correct Excel format
   - Check that well positions match between files
   - Verify sample names in FlowJo data contain well positions

2. **Missing Data**
   - Confirm all wells in FlowJo data have corresponding entries in maps
   - Check for typos in well positions
   - Ensure well positions use two-digit column numbers (A01 not A1)

3. **Export Issues**
   - Make sure you have write permissions for the save location
   - Close any open Excel files before saving

## Requirements

- Python 3.6+
- PyQt6
- pandas
- numpy
- scipy
- pyperclip

## Acknowledgments

Inspired by the legendary Dan Piraner
