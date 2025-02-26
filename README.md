# CASscanner - Python version with webui (dist folder)

## Overview
This repository contains scripts to scan detailed CAS (Consolidated Account Statement) shared via Karvy. The scripts analyze the CAS statement and extract fund details into separate data files.

## Python Files
The Python files in this repository are designed to handle the extraction, processing, and analysis of the CAS data. They perform the following tasks:
- Parse the CAS PDF files.
- Extract fund and transaction details.
- Store the data in a structured format for further analysis.

### List of Python Files
- `parser.py`: Parses the CAS PDF files and extracts relevant data.
- `processor.py`: Processes the extracted data and performs necessary calculations.
- `analyzer.py`: Analyzes the processed data and generates reports.

## Distribution Folder
The `dist` folder contains the distribution files for the project. This includes:
- Compiled executables for different platforms.
- Dependencies and libraries required to run the project.

## Usage
To use the Python scripts or the compiled executables, follow these steps:

1. **Clone the Repository**:
   ```bash
   git clone https://github.com/itsddpanda/CASscanner.git
   cd CASscanner
   ```

2. **Install Dependencies**:
   ```bash
   pip install -r requirements.txt
   ```

3. **Run the Scripts**:
   ```bash
   python app.py
   ```


## Contributions
Contributions are welcome! If you have any suggestions or improvements, feel free to fork the repository and submit a pull request.

## License
This project is licensed under the MIT License. See the LICENSE file for more details. It uses casparser and other opens source libraries, their licenses will be applicable also. 

