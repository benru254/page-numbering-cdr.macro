# Page Numbering Macro for CorelDRAW

## Overview
This VBA macro for CorelDRAW automates the process of inserting page numbers into a document. The macro allows users to specify font properties, positioning, and optional prefixes or suffixes for the page numbers.

## Features
- **Custom Font & Size**: Choose the font style and size for page numbers.
- **Positioning Options**: Align page numbers to the left, center, or right.
- **Prefix/Suffix Support**: Optionally add text before or after the page number.
- **Automatic Insertion**: Loops through all pages and inserts numbers dynamically.

## How It Works
1. Run the `ShowPageNumberForm` macro to open the user form.
2. Choose the desired font, font size, and positioning.
3. Optionally enter a prefix or suffix.
4. Click the insert button to apply page numbers across all pages.

## Installation
1. Open CorelDRAW.
2. Press `Alt + F11` to open the VBA editor.
3. Insert a new module and copy-paste the macro code.
4. Save and run the macro.

## Code Structure
- `ShowPageNumberForm()`: Displays the user form.
- `InsertPageNumbers(frm As Object)`: Retrieves user inputs and inserts page numbers.
- Uses `CreateArtisticText` to generate page numbers dynamically.

## Requirements
- CorelDRAW with VBA enabled.
- Basic understanding of VBA for modifications.

## License
This project is open-source under the MIT License.

## Author
Developed by **SMASH GRAPHICS AND SIGNS**. Contributions and improvements are welcome!

