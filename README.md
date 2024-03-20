# LC Sort: Library of Congress Call Number Sorter

## Overview
This application sorts .xlsx files containing Library of Congress (LoC) call numbers located in the first column (Column A). Designed to streamline the organization of collections cataloged with LoC call numbers, this tool ensures efficient sorting directly from an Excel spreadsheet, enhancing accessibility and manageability.

## Features
- **Excel Compatibility:** Works directly with .xlsx files to sort Library of Congress call numbers.
- **User-Friendly Interface:** Simple GUI prompts guide the user through file selection and provide clear instructions for use.
- **Error Handling:** Includes checks for file type and column header verification to prevent common errors.
- **Completion Notification:** Displays a message upon successful completion of the sorting process.
- **Flexibility:** Retains rows with 'N/A' or other specified placeholders, ensuring no data is lost in the sorting process.

## Prerequisites
- Windows operating system.
- No Python installation required for the standalone .exe version.

## Usage Instructions
1. **Start the Application:** Double-click on the .exe file to launch the application.
2. **Follow On-screen Instructions:** Read the initial instructions carefully to ensure your .xlsx file is properly formatted.
3. **Select Input File:** Use the file dialog window to choose the .xlsx file you wish to sort. The file must have 'Call Number' as the header in the first column.
4. **Processing:** The application will process the selected file. Rows with 'N/A' in the call number column are handled accordingly and not discarded.
5. **Completion:** Once processing is complete, a message box will notify you that the output file has been successfully saved. The output file is named after the original file with '_LCSortable' appended, saved in the same directory.

## Troubleshooting
- **File Type Error:** Ensure you select an .xlsx file. The application will alert you if a file of a different type is selected.
- **Column Header Error:** If the first column is not labeled 'Call Number', the application will notify you to correct the file and try again.
- **SmartScreen Warning:** Upon first launch, Windows Defender SmartScreen may prompt a warning. Click on "More info" and select "Run anyway" to proceed.

## Acknowledgments
This script was developed by Roger Peters with assistance from ChatGPT. The starting point for this script was the sortlc.py script originally developed for PALNI by lpmagnuson. The GitHub repository can be found here: https://github.com/PALNI/wmsinventory-lcsort/

The original script has been extended and adapted to include additional features, such as a graphical user interface for file selection and instructions, and to handle Excel files directly, enhancing usability and functionality.

Special thanks to lpmagnuson, the developer of the original sortlc.py script for the foundational work which provided a starting point for this project.
