# Homework Viewer

This is a homework viewer tool written in Python and Tkinter, which can read the homework content from Excel files for a specific date and display it in a GUI interface.

## Environment Requirements

- Python 3.6 or higher
- tkinter library
- openpyxl library
- datetime library

## Usage

1. Put the Excel files in the data directory under the py file root directory, the file name can be customized, but must end with .xlsx.
2. Run the py file, a GUI window will pop up.
3. Enter the date in the window, format as #m.#d, such as 1.29, 2.3, etc., and then click the view homework button.
4. The window will display the homework content for that date, the source is the parent directory or file name of the Excel file, depending on the value of the source_option parameter.
5. If you want to view the homework for other dates, you can repeat step 3.
6. If you want to exit the program, you can close the window.

## Parameter Description

- source_option: source option, 0 means the parent directory of the file, 1 means the file name, default is 1.
- directory: homework directory path, used to store Excel files, default is "work_tool\data".
- Other parameters, such as the window size, title, font, etc., can be modified according to personal preferences.
