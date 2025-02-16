# Excel Viewer & Data Entry Form

## Overview
This is a Python-based GUI application built with Tkinter that allows users to load an Excel file, view its data, and insert new entries into the file. The application provides a modern user interface with a sidebar for data entry, a dark mode toggle, and a treeview for displaying Excel data.

## Features
- **Load Excel Files**: Open `.xlsx` and `.xls` files to view data.
- **Data Entry Form**: Add new entries with Name, Age, Subscription status, and Employment status.
- **Dark Mode**: Toggle between dark and light themes.
- **Interactive UI**: Modern interface using Tkinter and ttk widgets.
- **Excel Data Handling**: Load and update Excel files using `openpyxl`.

## Installation
### Prerequisites
Ensure you have Python installed (Python 3.x recommended). You also need to install `openpyxl` for handling Excel files.

```sh
pip install openpyxl
```

### Clone the Repository
```sh
git clone https://github.com/yourusername/excel-viewer.git
cd excel-viewer
```

## Usage
Run the application using:
```sh
python app.py
```

### How to Use
1. **Load an Excel File**: Click "Load Excel File" and select an `.xlsx` file.
2. **View Data**: The loaded Excel file's content will appear in the treeview.
3. **Insert Data**: Enter details in the sidebar form and click "Insert" to add a new row.
4. **Toggle Dark Mode**: Use the "Dark Mode" toggle button to switch themes.

## Project Structure
```
├── app.py         # Main Python script
├── README.md      # Documentation
└── requirements.txt # List of dependencies (if needed)
```

## Dependencies
- `tkinter`
- `openpyxl`

## Contributing
Feel free to fork this repository and submit pull requests for improvements.

## License
This project is licensed under the MIT License.

