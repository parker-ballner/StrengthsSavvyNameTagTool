# Strengths Savvy Name Tag Tool

This is a simple data parsing tool used to make Strengths Savvy name tags

## Prerequisites:

### 1: Download python3 from [python.org](https://www.python.org) (or if using macOs, install Homebrew and install python3 using Homebrew)

### 2: Install dependencies:
```
pip install pandas
pip install pdfkit
pip install imgkit
```

### 3: Install wkhtmltopdf

Debian/Ubuntu:
```
sudo apt-get install wkhtmltopdf
```

macOs and Windows:
Download and install the correct version from the [wkhtmltopdf webiste](http://wkhtmltopdf.org/downloads.html)

### 4: Update the wkhtmltopdf_path.txt file with the proper directory for the wkhtmltopdf binary/executable that you downloaded. **NOTE:** When adding the directory to the txt file, be sure to used forward slashes (/) and not back slashes. (The program won't like it). The path needs to be specified or the program won't produce the pdf, just the html.

Typical Locations:

- macOs: /Applications/wkhtmltopdf.app/Contents/MacOS/wkhtmltopdf

- Windows: C:/Program Files/wkhtmltopdf/bin/wkhtmltopdf.exe

### 5: Run the program!

1. Clean out output folder
2. Clean out input folder and add new files (check input folder readme)
3. `python3 ss_nametag_tool.py '{name_of_xlsx_file_in_input_directory.xlsx}' '{name_of_logo_file_in_input_directory.png}'`

### Bonus: Digital Name Tags

Create zoom backgrounds for each user by adding the 'digital' flag to the end of the command
`python3 ss_nametag_tool.py '{name_of_xlsx_file_in_input_directory.xlsx}' '{name_of_logo_file_in_input_directory.png}' digital`