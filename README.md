
# E2D Report Generator (LOCAL)

The **E2D Report Generator** is a web-based tool that allows users to generate reports based on data from an Excel file. The reports are generated in Word document format using customizable templates.

## Features

- **Easy Report Generation:** Upload an Excel file and select a template to generate reports automatically.
- **Customizable Templates:** Choose between different border templates (1st Border or 2nd Border).
- **Automated Document Creation:** Automatically populate Word documents with data from the Excel file.
- **ZIP File Download:** Download all generated reports in a ZIP file.

## How to Use

### Upload Your Excel File:
- Navigate to the system’s interface.
- Choose your Excel file that contains the necessary data for report generation.

### Select Template:
- Choose the desired border template (either "1st Border" or "2nd Border").

### Generate Reports:
- Click on the **Generate Reports** button. The system will process the data from the Excel file and generate reports based on the selected template.

### Download Reports:
- After the reports are generated, a ZIP file will be provided for download containing all the Word documents.

## Installation

To run the project locally:

1. Clone the repository:

   ```bash
   git clone https://github.com/JB-VERGARA/e2d-report-generator.git
   ```

2. Install dependencies:

   ```bash
   npm install
   ```

3. Start the server:

   ```bash
   node server.js
   ```

4. Visit `http://localhost:3000` in your browser to use the system.
